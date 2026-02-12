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

            # 3. 본문(section0.xml) 파싱 및 변환
            section_xml = z.read('Contents/section0.xml')
            root = etree.fromstring(section_xml)

            # 4. 양식 템플릿 정보 보존
            self._save_template_info(z, root)

        # 5. 문서 트리 순회하며 마크다운 생성
        md_lines = self._process_section(root)
        return '\n'.join(md_lines)

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

    def _process_section(self, root):
        """섹션 루트 아래의 최상위 문단들을 순회"""
        lines = []
        for para in root.findall('hp:p', NS):
            para_lines = self._process_paragraph(para, top_level=True)
            lines.extend(para_lines)
        return lines

    def _process_paragraph(self, para, top_level=False):
        """하나의 hp:p를 처리"""
        lines = []
        para_pr_id = para.get('paraPrIDRef', '0')

        # 제목 감지
        heading_level = None
        if self.style_map:
            heading_level = self.style_map.get_heading_level(para_pr_id)

        # 테이블 감지 - 문단 내 테이블이 있으면 테이블로 처리
        tables = para.findall('.//hp:tbl', NS)

        # 이미지 감지
        pics = para.findall('.//hp:pic', NS)

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
        elif text.strip():
            if heading_level:
                lines.append(f"\n{'#' * heading_level} {text.strip()}\n")
            else:
                lines.append(text.strip())
        elif top_level:
            lines.append('')  # 빈 줄 보존

        return lines

    def _extract_paragraph_text(self, para):
        """문단에서 인라인 텍스트 추출 (서식 포함)"""
        parts = []
        for run in para.findall('hp:run', NS):
            char_pr_id = run.get('charPrIDRef', '0')
            fmt = {}
            if self.style_map:
                fmt = self.style_map.get_char_format(char_pr_id)

            for child in run:
                tag = etree.QName(child.tag).localname
                if tag == 't':
                    raw_text = child.text or ''
                    # 내부에 lineBreak 같은 요소가 있을 수 있음
                    for sub in child:
                        sub_tag = etree.QName(sub.tag).localname
                        if sub_tag == 'lineBreak':
                            raw_text += '\n'
                        if sub.tail:
                            raw_text += sub.tail

                    if raw_text:
                        formatted = self._apply_format(raw_text, fmt)
                        parts.append(formatted)

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
