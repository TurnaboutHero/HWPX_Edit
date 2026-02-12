"""
hwpx_to_md.py - HWPX to Markdown 변환기
HWPX 파일(ZIP 내부 XML)을 파싱하여 Markdown으로 변환합니다.
"""
import os
import sys
import zipfile
import argparse
import json
from lxml import etree


# HWPX XML 네임스페이스
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
}


class HwpxStyleMap:
    """header.xml에서 스타일 정보를 추출하여 제목 레벨 등을 판별"""

    def __init__(self, header_xml_bytes):
        self.root = etree.fromstring(header_xml_bytes)
        self.outline_levels = {}   # paraPrIDRef -> outline_level (0-based)
        self.char_props = {}       # charPrID -> {bold, italic, underline, ...}
        self._parse_outline_levels()
        self._parse_char_properties()

    def _parse_outline_levels(self):
        for para_pr in self.root.findall('.//hh:paraPr', NS):
            pr_id = para_pr.get('id')
            heading = para_pr.find('.//hh:heading', NS)
            if heading is not None and heading.get('type') == 'OUTLINE':
                level_str = heading.get('level')
                if level_str is not None:
                    self.outline_levels[pr_id] = int(level_str)

    def _parse_char_properties(self):
        for char_pr in self.root.findall('.//hh:charPr', NS):
            cp_id = char_pr.get('id')
            props = {
                'bold': char_pr.find('hh:bold', NS) is not None,
                'italic': char_pr.find('hh:italic', NS) is not None,
                'underline': False,
                'strikeout': False,
            }
            ul = char_pr.find('hh:underline', NS)
            if ul is not None and ul.get('type', 'NONE') != 'NONE':
                props['underline'] = True
            st = char_pr.find('hh:strikeout', NS)
            if st is not None and st.get('shape', 'NONE') != 'NONE':
                props['strikeout'] = True
            self.char_props[cp_id] = props

    def get_heading_level(self, para_pr_id):
        """paraPrIDRef로 제목 레벨(1-based) 반환. 제목 아니면 None."""
        level = self.outline_levels.get(str(para_pr_id))
        if level is not None:
            return level + 1  # 0-based → 1-based
        return None

    def get_char_format(self, char_pr_id):
        return self.char_props.get(str(char_pr_id), {})


class HwpxToMarkdown:
    """HWPX 파일을 Markdown으로 변환"""

    def __init__(self, hwpx_path, output_dir=None, extract_images=True):
        self.hwpx_path = hwpx_path
        self.output_dir = output_dir or os.path.dirname(hwpx_path) or '.'
        self.extract_images = extract_images
        self.images_dir = os.path.join(self.output_dir, 'images')
        self.style_map = None
        self.image_map = {}  # binaryItemIDRef -> extracted_filename
        self.template_info = {}  # 양식 보존용 메타데이터
        self.footnotes = []  # (ref_num, text) 튜플 리스트
        self.endnotes = []  # (ref_num, text) 튜플 리스트
        self.footnote_counter = 0
        self.endnote_counter = 0

    def convert(self):
        """메인 변환 함수. 마크다운 문자열 반환."""
        with zipfile.ZipFile(self.hwpx_path, 'r') as z:
            # 1. 헤더(스타일 정보) 파싱
            if 'Contents/header.xml' in z.namelist():
                header_bytes = z.read('Contents/header.xml')
                self.style_map = HwpxStyleMap(header_bytes)

            # 2. 이미지 추출
            if self.extract_images:
                self._extract_images(z)

            # 3. 다중 섹션 찾기 및 정렬
            section_files = self._find_section_files(z)

            # 4. 각 섹션 변환 및 병합
            all_md_lines = []
            first_section = True

            for section_file in section_files:
                section_xml = z.read(section_file)
                root = etree.fromstring(section_xml)

                # 첫 섹션에서 양식 정보 저장
                if first_section:
                    self._save_template_info(z, root)
                    first_section = False

                # 섹션 변환
                md_lines = self._process_section(root)
                all_md_lines.extend(md_lines)

                # 섹션 구분자 추가 (마지막 섹션 제외)
                if section_file != section_files[-1]:
                    all_md_lines.append('')
                    all_md_lines.append('---')
                    all_md_lines.append('')

        # 5. 각주/미주 정의 추가
        if self.footnotes or self.endnotes:
            all_md_lines.append('')
            all_md_lines.append('')

            if self.footnotes:
                for ref_num, text in self.footnotes:
                    all_md_lines.append(f"[^{ref_num}]: {text}")

            if self.endnotes:
                for ref_num, text in self.endnotes:
                    all_md_lines.append(f"[^e{ref_num}]: {text}")

        return '\n'.join(all_md_lines)

    def _extract_images(self, z):
        """BinData 폴더의 이미지를 추출"""
        os.makedirs(self.images_dir, exist_ok=True)
        for name in z.namelist():
            if name.startswith('BinData/'):
                basename = os.path.basename(name)
                if not basename:
                    continue
                # binaryItemIDRef는 확장자 없는 이름 (e.g., "image1")
                ref_id = os.path.splitext(basename)[0]
                ext = os.path.splitext(basename)[1]

                # BMP → PNG 변환 (파일 크기 절약)
                out_name = f"{ref_id}.png" if ext.lower() == '.bmp' else basename
                out_path = os.path.join(self.images_dir, out_name)

                img_data = z.read(name)
                if ext.lower() == '.bmp':
                    try:
                        from PIL import Image
                        import io
                        img = Image.open(io.BytesIO(img_data))
                        img.save(out_path, 'PNG')
                    except ImportError:
                        # Pillow 없으면 BMP 그대로 저장
                        out_name = basename
                        out_path = os.path.join(self.images_dir, out_name)
                        with open(out_path, 'wb') as f:
                            f.write(img_data)
                else:
                    with open(out_path, 'wb') as f:
                        f.write(img_data)

                self.image_map[ref_id] = out_name

    def _find_section_files(self, z):
        """ZIP 내부의 모든 section*.xml 파일을 찾아 정렬하여 반환"""
        import re
        section_files = []
        pattern = re.compile(r'^Contents/section(\d+)\.xml$')

        for name in z.namelist():
            match = pattern.match(name)
            if match:
                section_num = int(match.group(1))
                section_files.append((section_num, name))

        # 번호 순서대로 정렬
        section_files.sort(key=lambda x: x[0])
        return [name for _, name in section_files]

    def _save_template_info(self, z, section_root):
        """양식 정보를 JSON으로 보존 (나중에 hwpx 복원 시 사용)"""
        info = {
            'source_file': os.path.basename(self.hwpx_path),
            'files_in_hwpx': z.namelist(),
            'images': list(self.image_map.keys()),
        }

        # 페이지 설정 추출
        page_pr = section_root.find('.//hp:pagePr', NS)
        if page_pr is not None:
            info['page'] = {
                'width': page_pr.get('width'),
                'height': page_pr.get('height'),
                'landscape': page_pr.get('landscape'),
            }
            margin = page_pr.find('hp:margin', NS)
            if margin is not None:
                info['page']['margins'] = dict(margin.attrib)

        self.template_info = info

        # JSON으로 저장
        info_path = os.path.join(self.output_dir, 'template_info.json')
        with open(info_path, 'w', encoding='utf-8') as f:
            json.dump(info, f, ensure_ascii=False, indent=2)

    def _collect_footnotes_endnotes(self, para):
        """문단에서 각주/미주를 수집하고 참조 번호 부여"""
        # 각주 수집
        for footnote in para.findall('.//hp:footnote', NS):
            self.footnote_counter += 1
            ref_num = self.footnote_counter

            # 각주 텍스트 추출
            footnote_texts = []
            for sub_list in footnote.findall('.//hp:subList', NS):
                for sub_para in sub_list.findall('.//hp:p', NS):
                    text = self._extract_paragraph_text(sub_para)
                    if text.strip():
                        footnote_texts.append(text.strip())

            if footnote_texts:
                footnote_text = ' '.join(footnote_texts)
                self.footnotes.append((ref_num, footnote_text))

        # 미주 수집
        for endnote in para.findall('.//hp:endnote', NS):
            self.endnote_counter += 1
            ref_num = self.endnote_counter

            # 미주 텍스트 추출
            endnote_texts = []
            for sub_list in endnote.findall('.//hp:subList', NS):
                for sub_para in sub_list.findall('.//hp:p', NS):
                    text = self._extract_paragraph_text(sub_para)
                    if text.strip():
                        endnote_texts.append(text.strip())

            if endnote_texts:
                endnote_text = ' '.join(endnote_texts)
                self.endnotes.append((ref_num, endnote_text))

    def _process_section(self, root):
        """섹션 루트 아래의 최상위 문단들을 순회"""
        lines = []

        # 1. 머리글 추출
        headers = root.findall('.//hp:header', NS)
        for header in headers:
            header_text = self._extract_header_footer_text(header)
            if header_text:
                lines.append(f"<!-- 머리글: {header_text} -->")
                lines.append('')

        # 2. 본문 문단 처리
        for para in root.findall('hp:p', NS):
            para_lines = self._process_paragraph(para, top_level=True)
            lines.extend(para_lines)

        # 3. 꼬리글 추출
        footers = root.findall('.//hp:footer', NS)
        for footer in footers:
            footer_text = self._extract_header_footer_text(footer)
            if footer_text:
                lines.append('')
                lines.append(f"<!-- 꼬리글: {footer_text} -->")

        return lines

    def _process_paragraph(self, para, top_level=False):
        """하나의 hp:p를 처리"""
        lines = []
        para_pr_id = para.get('paraPrIDRef', '0')

        # 제목 감지
        heading_level = None
        if self.style_map:
            heading_level = self.style_map.get_heading_level(para_pr_id)

        # 각주/미주 수집
        self._collect_footnotes_endnotes(para)

        # 테이블 감지 - 문단 내 테이블이 있으면 테이블로 처리
        tables = para.findall('.//hp:tbl', NS)

        # 이미지 감지
        pics = para.findall('.//hp:pic', NS)

        # 수식 감지
        equations = para.findall('.//hp:equation', NS)

        # 양식 개체 감지
        form_elements = []
        form_elements.extend(para.findall('.//hp:checkBtn', NS))
        form_elements.extend(para.findall('.//hp:radioBtn', NS))
        form_elements.extend(para.findall('.//hp:comboBox', NS))
        form_elements.extend(para.findall('.//hp:btn', NS))
        form_elements.extend(para.findall('.//hp:edit', NS))

        # TextArt (글맵시) 감지
        textarts = para.findall('.//hp:textart', NS)

        # OLE 개체 감지
        oles = para.findall('.//hp:ole', NS)

        # 다단 레이아웃 감지
        colprs = para.findall('.//hp:colPr', NS)

        # 도형/글상자 감지
        shape_tags = ['hp:rect', 'hp:ellipse', 'hp:arc', 'hp:polygon', 'hp:curve', 'hp:connectLine', 'hp:container']
        shapes_with_text = []
        for shape_tag in shape_tags:
            shapes = para.findall(f'.//{shape_tag}', NS)
            for shape in shapes:
                shape_text = self._extract_shape_text(shape)
                if shape_text:
                    shapes_with_text.append(shape_text)

        # 텍스트 추출
        text = self._extract_paragraph_text(para)

        # 테이블 처리
        if tables:
            for tbl in tables:
                tbl_lines = self._process_table(tbl)
                lines.extend(tbl_lines)

            # 테이블 외 텍스트가 있으면 추가
            if text.strip():
                if heading_level:
                    lines.insert(0, f"\n{'#' * heading_level} {text.strip()}\n")
                else:
                    lines.insert(0, text.strip())
        elif pics:
            # 이미지가 있는 문단
            for pic in pics:
                img_line = self._process_image(pic)
                if img_line:
                    lines.append(img_line)
            if text.strip():
                lines.insert(0, text.strip())
        elif equations:
            # 수식이 있는 문단
            if text.strip():
                lines.append(text.strip())
            for eq in equations:
                eq_line = self._process_equation(eq)
                if eq_line:
                    lines.append(eq_line)
        elif form_elements:
            # 양식 개체가 있는 문단
            if text.strip():
                lines.append(text.strip())
            for form_elem in form_elements:
                form_line = self._process_form_element(form_elem)
                if form_line:
                    lines.append(form_line)
        elif text.strip():
            if heading_level:
                lines.append(f"\n{'#' * heading_level} {text.strip()}\n")
            else:
                lines.append(text.strip())
        elif top_level:
            lines.append('')  # 빈 줄 보존

        # 다단 레이아웃 메타데이터 삽입
        for colpr in colprs:
            colpr_line = self._process_colpr(colpr)
            if colpr_line:
                lines.append(colpr_line)

        # TextArt (글맵시) 처리
        for textart in textarts:
            ta_line = self._process_textart(textart)
            if ta_line:
                lines.append(ta_line)

        # OLE 개체 처리
        for ole in oles:
            ole_line = self._process_ole(ole)
            if ole_line:
                lines.append(ole_line)

        # 도형/글상자 텍스트 추가
        for shape_text in shapes_with_text:
            lines.append(f"> [글상자] {shape_text}")

        return lines

    def _extract_paragraph_text(self, para):
        """문단에서 인라인 텍스트 추출 (서식 포함)"""
        parts = []

        # 각주/미주 요소를 미리 수집하여 인덱스 매핑 생성
        all_footnotes = para.findall('.//hp:footnote', NS)
        all_endnotes = para.findall('.//hp:endnote', NS)

        # 현재 문단에서 각주/미주의 시작 번호 계산
        footnote_start = self.footnote_counter - len(all_footnotes) + 1
        endnote_start = self.endnote_counter - len(all_endnotes) + 1

        # 하이퍼링크 상태 추적
        hyperlink_url = None

        for run in para.findall('hp:run', NS):
            char_pr_id = run.get('charPrIDRef', '0')
            fmt = {}
            if self.style_map:
                fmt = self.style_map.get_char_format(char_pr_id)

            for child in run:
                tag = etree.QName(child.tag).localname
                if tag == 't':
                    raw_text = child.text or ''
                    # 내부에 lineBreak, 변경 추적 등의 요소가 있을 수 있음
                    for sub in child:
                        sub_tag = etree.QName(sub.tag).localname
                        if sub_tag == 'lineBreak':
                            raw_text += '\n'
                            if sub.tail:
                                raw_text += sub.tail
                            continue
                        elif sub_tag == 'deleteBegin':
                            # deleteBegin~deleteEnd 사이 텍스트는 취소선
                            del_text = sub.tail or ''
                            if del_text:
                                # 앞뒤 공백을 ~~ 바깥으로 이동 (Markdown 호환)
                                leading = del_text[:len(del_text) - len(del_text.lstrip())]
                                trailing = del_text[len(del_text.rstrip()):]
                                core = del_text.strip()
                                if core:
                                    raw_text += f'{leading}~~{core}~~{trailing}'
                                else:
                                    raw_text += del_text
                            continue  # tail 이미 처리됨
                        elif sub_tag == 'deleteEnd':
                            if sub.tail:
                                raw_text += sub.tail
                            continue
                        elif sub_tag == 'insertBegin':
                            # insertBegin~insertEnd 사이 텍스트는 그대로 출력
                            if sub.tail:
                                raw_text += sub.tail
                            continue
                        elif sub_tag == 'insertEnd':
                            if sub.tail:
                                raw_text += sub.tail
                            continue
                        else:
                            if sub.tail:
                                raw_text += sub.tail

                    if raw_text:
                        formatted = self._apply_format(raw_text, fmt)
                        # 하이퍼링크 컨텍스트 내 텍스트면 링크로 감싸기
                        if hyperlink_url:
                            parts.append(f"[{formatted}]({hyperlink_url})")
                            hyperlink_url = None  # URL 소비
                        else:
                            parts.append(formatted)

                elif tag == 'ctrl':
                    # 각주/미주 참조 삽입
                    footnote = child.find('hp:footnote', NS)
                    endnote = child.find('hp:endnote', NS)

                    if footnote is not None and footnote in all_footnotes:
                        idx = all_footnotes.index(footnote)
                        ref_num = footnote_start + idx
                        parts.append(f"[^{ref_num}]")
                    elif endnote is not None and endnote in all_endnotes:
                        idx = all_endnotes.index(endnote)
                        ref_num = endnote_start + idx
                        parts.append(f"[^e{ref_num}]")

                elif tag == 'dutmal':
                    # 덧말 (Ruby Text): <ruby>본말<rt>닷말</rt></ruby>
                    main_el = child.find('hp:mainText', NS)
                    sub_el = child.find('hp:subText', NS)
                    main_text = main_el.text if main_el is not None and main_el.text else ''
                    sub_text = sub_el.text if sub_el is not None and sub_el.text else ''
                    if main_text:
                        parts.append(f"<ruby>{main_text}<rt>{sub_text}</rt></ruby>")

                elif tag == 'fieldBegin':
                    # 하이퍼링크 필드 감지
                    field_type = child.get('type', '')
                    if field_type == 'HYPERLINK':
                        # URL을 parameters에서 추출
                        params = child.find('.//hp:parameters', NS)
                        if params is not None:
                            for sp in params.findall('hp:stringParam', NS):
                                if sp.get('name') == 'url' and sp.text:
                                    hyperlink_url = sp.text
                                    break
                        # 대체: href 속성에서도 시도
                        if not hyperlink_url:
                            hyperlink_url = child.get('href', '') or None

                elif tag == 'fieldEnd':
                    # 하이퍼링크 필드 종료
                    hyperlink_url = None

        return ''.join(parts)

    def _apply_format(self, text, fmt):
        """인라인 서식 적용"""
        if not fmt or not text.strip():
            return text

        result = text
        if fmt.get('bold') and fmt.get('italic'):
            result = f"***{result}***"
        elif fmt.get('bold'):
            result = f"**{result}**"
        elif fmt.get('italic'):
            result = f"*{result}*"

        if fmt.get('strikeout'):
            result = f"~~{result}~~"

        return result

    def _process_table(self, tbl):
        """hp:tbl을 마크다운 테이블로 변환"""
        lines = []
        row_cnt = int(tbl.get('rowCnt', 0))
        col_cnt = int(tbl.get('colCnt', 0))

        if row_cnt == 0 or col_cnt == 0:
            return lines

        # 셀 데이터를 2D 그리드로 수집
        grid = [['' for _ in range(col_cnt)] for _ in range(row_cnt)]
        occupied = [[False for _ in range(col_cnt)] for _ in range(row_cnt)]

        for tr in tbl.findall('.//hp:tr', NS):
            for tc in tr.findall('hp:tc', NS):
                addr = tc.find('hp:cellAddr', NS)
                span = tc.find('hp:cellSpan', NS)

                if addr is None:
                    continue

                col = int(addr.get('colAddr', 0))
                row = int(addr.get('rowAddr', 0))
                colspan = int(span.get('colSpan', 1)) if span is not None else 1
                rowspan = int(span.get('rowSpan', 1)) if span is not None else 1

                # 셀 내 텍스트 추출
                cell_texts = []
                for sub_para in tc.findall('.//hp:p', NS):
                    p_text = self._extract_paragraph_text(sub_para)
                    if p_text.strip():
                        cell_texts.append(p_text.strip())

                cell_text = ' '.join(cell_texts)
                # 마크다운 테이블에서 파이프 이스케이프
                cell_text = cell_text.replace('|', '\\|')
                # 줄바꿈은 <br>로
                cell_text = cell_text.replace('\n', '<br>')

                # 그리드에 채우기
                if row < row_cnt and col < col_cnt:
                    grid[row][col] = cell_text

                # 병합 셀 표시
                for r in range(row, min(row + rowspan, row_cnt)):
                    for c in range(col, min(col + colspan, col_cnt)):
                        occupied[r][c] = True
                        if r == row and c == col:
                            continue
                        # 병합된 셀은 빈 문자열로 (이미 초기화됨)

        # 마크다운 테이블 생성
        if row_cnt == 0:
            return lines

        lines.append('')  # 테이블 전 빈줄

        # 1x1 테이블은 인용문으로 변환
        if row_cnt == 1 and col_cnt == 1:
            content = grid[0][0]
            if content:
                lines.append(f"> {content}")
                lines.append('')
                return lines

        # 일반 테이블
        # 헤더 행
        header = '| ' + ' | '.join(grid[0]) + ' |'
        separator = '| ' + ' | '.join(['---'] * col_cnt) + ' |'
        lines.append(header)
        lines.append(separator)

        # 데이터 행
        for r in range(1, row_cnt):
            row_str = '| ' + ' | '.join(grid[r]) + ' |'
            lines.append(row_str)

        lines.append('')  # 테이블 후 빈줄
        return lines

    def _process_image(self, pic):
        """hp:pic을 마크다운 이미지로 변환"""
        img_el = pic.find('.//hc:img', NS)
        if img_el is None:
            return None

        ref_id = img_el.get('binaryItemIDRef', '')
        if ref_id in self.image_map:
            filename = self.image_map[ref_id]
            return f"\n![{ref_id}](images/{filename})\n"
        return f"\n![{ref_id}](images/{ref_id})\n"

    def _extract_header_footer_text(self, element):
        """머리글/꼬리글에서 텍스트 추출"""
        texts = []
        for para in element.findall('.//hp:p', NS):
            para_text = self._extract_paragraph_text(para)
            if para_text.strip():
                texts.append(para_text.strip())
        return ' '.join(texts)

    def _extract_shape_text(self, shape):
        """도형/글상자 안의 텍스트 추출"""
        draw_text = shape.find('.//hp:drawText', NS)
        if draw_text is None:
            return None

        texts = []
        for para in draw_text.findall('.//hp:p', NS):
            para_text = self._extract_paragraph_text(para)
            if para_text.strip():
                texts.append(para_text.strip())

        return ' '.join(texts) if texts else None

    def _process_textart(self, textart):
        """TextArt (글맵시)를 마크다운으로 변환"""
        text = textart.get('text', '')
        if not text:
            return None
        # 특수 유니코드 CR/LF (U+240D, U+240A) 및 실제 \r\n을 공백으로
        text = text.replace('\u240d\u240a', ' ')
        text = text.replace('\u240d', ' ')
        text = text.replace('\u240a', ' ')
        text = text.replace('\r\n', ' ')
        text = text.replace('\r', ' ')
        text = text.replace('\n', ' ')
        text = text.strip()
        if text:
            return f"> [글맵시] {text}"
        return None

    def _process_ole(self, ole):
        """OLE 개체를 마크다운 주석으로 변환"""
        binary_ref = ole.get('binaryItemIDRef', '')
        # shapeComment에서 개체 정보 추출
        comment_el = ole.find('hp:shapeComment', NS)
        comment = ''
        if comment_el is not None and comment_el.text:
            comment = comment_el.text.strip().replace('\n', ' ').replace('\r', '')
        if comment:
            return f"<!-- OLE: {comment} ({binary_ref}) -->"
        obj_type = ole.get('objectType', 'EMBEDDED')
        return f"<!-- OLE: {obj_type} ({binary_ref}) -->"

    def _process_colpr(self, colpr):
        """다단 레이아웃 메타데이터를 마크다운 주석으로 변환"""
        col_count = int(colpr.get('colCount', '1'))
        if col_count > 1:
            return f"<!-- [{col_count}단 레이아웃] -->"
        return None

    def _process_equation(self, equation):
        """수식을 마크다운으로 변환"""
        script = equation.find('.//hp:script', NS)
        if script is not None and script.text:
            # 한글 수식 스크립트를 그대로 유지
            return f"\n$$\n{script.text.strip()}\n$$\n"
        return None

    def _process_form_element(self, element):
        """양식 개체를 마크다운으로 변환"""
        tag = etree.QName(element.tag).localname

        caption = element.get('caption', '').strip()
        name = element.get('name', '').strip()
        value = element.get('value', 'UNCHECKED')

        # caption이 비어있으면 name 사용
        display_text = caption if caption else name

        if tag == 'checkBtn':
            # 체크박스: [ ] 또는 [x]
            checked = 'x' if value == 'CHECKED' else ' '
            return f"[{checked}] {display_text}"

        elif tag == 'radioBtn':
            # 라디오 버튼: ( ) 또는 (o)
            checked = 'o' if value == 'CHECKED' else ' '
            return f"({checked}) {display_text}"

        elif tag == 'comboBox':
            # 콤보박스: [콤보: name] 또는 [콤보: name = value]
            selected = element.get('selectedValue', '').strip()
            if selected:
                return f"[콤보: {name} = {selected}]"
            else:
                return f"[콤보: {name}]"

        elif tag == 'btn':
            # 버튼: [버튼: caption]
            return f"[버튼: {display_text}]"

        elif tag == 'edit':
            # 입력란: [입력란: name]
            return f"[입력란: {name}]"

        return None


def convert_hwpx_to_md(hwpx_path, output_path=None, extract_images=True):
    """hwpx 파일을 마크다운으로 변환하는 편의 함수"""
    if output_path is None:
        base = os.path.splitext(hwpx_path)[0]
        output_path = base + '.md'

    output_dir = os.path.dirname(output_path) or '.'

    converter = HwpxToMarkdown(hwpx_path, output_dir=output_dir, extract_images=extract_images)
    md_content = converter.convert()

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"변환 완료: {output_path}")
    if extract_images and converter.image_map:
        print(f"이미지 {len(converter.image_map)}개 추출: {converter.images_dir}")
    if converter.template_info:
        print(f"양식 정보 저장: {os.path.join(output_dir, 'template_info.json')}")

    return output_path


def main():
    parser = argparse.ArgumentParser(description='HWPX → Markdown 변환기')
    parser.add_argument('input', help='입력 HWPX 파일 경로')
    parser.add_argument('-o', '--output', help='출력 마크다운 파일 경로')
    parser.add_argument('--no-images', action='store_true', help='이미지 추출 안 함')
    args = parser.parse_args()

    convert_hwpx_to_md(args.input, args.output, extract_images=not args.no_images)


if __name__ == '__main__':
    main()
