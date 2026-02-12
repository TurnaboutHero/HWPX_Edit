"""
smart_replace.py - 원본 HWPX 구조 보존 + 마크다운 텍스트 반영

원본 HWPX의 바이트를 그대로 보존하면서, 편집된 마크다운에서 변경된
테이블 셀 텍스트만 원본 XML 문자열에 직접 치환합니다.

lxml은 분석(파싱)에만 사용하고, 직렬화는 하지 않습니다.
이를 통해 CRLF, 속성 순서, 네임스페이스 등이 원본과 100% 동일하게 보존됩니다.

주요 장점:
  - 표 서식 완벽 보존 (셀 병합, 너비, 높이, 테두리, 배경색)
  - 원본 XML 바이트 수준 보존 (lxml 직렬화 우회)
  - pypandoc-hwpx 버그 우회

제한사항:
  - 테이블 셀 텍스트만 교체 가능 (문단 텍스트는 원본 유지)
  - 구조 변경(행/열 추가/삭제)은 반영 불가
"""
import os
import re
import sys
import argparse
import zipfile
import io
import difflib
from lxml import etree


# HWPX XML 네임스페이스 — 2011 (한컴) / 2024 (OWPML 표준) 자동 감지
NS_2011 = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

NS_2024 = {
    'hp': 'http://www.owpml.org/owpml/2024/paragraph',
    'hs': 'http://www.owpml.org/owpml/2024/body',
    'hc': 'http://www.owpml.org/owpml/2024/core',
    'hh': 'http://www.owpml.org/owpml/2024/head',
}

# 기본값 (2011) — smart_replace() 시 자동 감지하여 교체
NS = dict(NS_2011)


def detect_namespace_version(xml_bytes):
    """XML 바이트에서 네임스페이스 버전 감지 (2011 vs 2024)"""
    snippet = xml_bytes[:2000] if isinstance(xml_bytes, bytes) else xml_bytes.encode()[:2000]
    if b'owpml.org/owpml/2024' in snippet:
        return '2024'
    return '2011'


def detect_close_tag(raw_xml):
    """raw XML 문자열에서 텍스트 태그 닫기 패턴을 감지.

    HWPX 2011은 </hp:t>, OWPML 2024는 </p:t> 또는 </owpml:t> 등
    다양한 접두사를 사용할 수 있으므로, 실제 XML에서 사용되는 패턴을 감지.

    Args:
        raw_xml: section XML 문자열

    Returns:
        str: 감지된 닫기 태그 (예: '</hp:t>'). 감지 실패 시 '</hp:t>' 기본값.
    """
    m = re.search(r'</[\w]+:t>', raw_xml)
    if m:
        return m.group(0)
    return '</hp:t>'


# ============================================================
# 마크다운 파서
# ============================================================

def parse_markdown_tables(md_text):
    """마크다운에서 테이블만 순서대로 추출 (인용문 포함).

    Returns:
        list of dict: {'type': 'table'|'quote', 'cells': 2D list}
    """
    tables = []
    lines = md_text.split('\n')
    i = 0

    while i < len(lines):
        line = lines[i]

        # 인용문 (hwpx_to_md에서 1×1 테이블을 인용문으로 변환)
        if line.startswith('> '):
            tables.append({
                'type': 'quote',
                'cells': [[line[2:].strip()]],
            })
            i += 1
            continue

        # 테이블 감지
        if '|' in line:
            next_i = i + 1
            if next_i < len(lines) and re.match(r'^\|[\s\-:|]+\|$', lines[next_i].strip()):
                table_lines = []
                while i < len(lines) and lines[i].strip() and '|' in lines[i]:
                    table_lines.append(lines[i])
                    i += 1
                cells = _parse_table_lines(table_lines)
                tables.append({
                    'type': 'table',
                    'cells': cells,
                })
                continue

        i += 1

    return tables


def _parse_table_lines(table_lines):
    """마크다운 테이블 행들을 2D 리스트로 변환"""
    cells = []
    for idx, line in enumerate(table_lines):
        if idx == 1 and re.match(r'^\|[\s\-:|]+\|$', line.strip()):
            continue
        stripped = line.strip()
        if stripped.startswith('|'):
            stripped = stripped[1:]
        if stripped.endswith('|'):
            stripped = stripped[:-1]
        row = [c.strip() for c in stripped.split('|')]
        cells.append(row)
    return cells


# ============================================================
# XML 분석 (lxml — 읽기 전용, 직렬화 안 함)
# ============================================================

def extract_xml_tables(section_root):
    """section0.xml에서 테이블 정보 추출 (인용문=1×1 테이블 포함).

    hwpx_to_md.py와 동일한 순서로 순회하여 마크다운 테이블과 1:1 매칭.

    Returns:
        list of dict: {'type', 'row_cnt', 'col_cnt', 'cells': 2D list}
    """
    tables = []
    for para in section_root.findall('hp:p', NS):
        for tbl in para.findall('.//hp:tbl', NS):
            row_cnt = int(tbl.get('rowCnt', 0))
            col_cnt = int(tbl.get('colCnt', 0))
            cells = _get_table_cells(tbl, row_cnt, col_cnt)
            is_quote = (row_cnt == 1 and col_cnt == 1)
            tables.append({
                'type': 'quote' if is_quote else 'table',
                'row_cnt': row_cnt,
                'col_cnt': col_cnt,
                'cells': cells,
            })
    return tables


def _get_para_text(para):
    """hp:p에서 순수 텍스트 추출"""
    parts = []
    for run in para.findall('hp:run', NS):
        for child in run:
            tag = etree.QName(child.tag).localname
            if tag == 't':
                text = child.text or ''
                for sub in child:
                    sub_tag = etree.QName(sub.tag).localname
                    if sub_tag == 'lineBreak':
                        text += '\n'
                    if sub.tail:
                        text += sub.tail
                parts.append(text)
    return ''.join(parts)


def _get_table_cells(tbl, row_cnt, col_cnt):
    """hp:tbl에서 셀 텍스트를 2D 리스트로 추출"""
    grid = [['' for _ in range(col_cnt)] for _ in range(row_cnt)]

    for tr in tbl.findall('.//hp:tr', NS):
        for tc in tr.findall('hp:tc', NS):
            addr = tc.find('hp:cellAddr', NS)
            if addr is None:
                continue
            col = int(addr.get('colAddr', 0))
            row = int(addr.get('rowAddr', 0))

            cell_texts = []
            for p in tc.findall('.//hp:p', NS):
                text = _get_para_text(p)
                if text.strip():
                    cell_texts.append(text.strip())

            if row < row_cnt and col < col_cnt:
                grid[row][col] = ' '.join(cell_texts)

    return grid


# ============================================================
# 텍스트 정규화 & 비교
# ============================================================

def _strip_md_format(text):
    """마크다운 인라인 서식 제거"""
    text = re.sub(r'\*{3}(.+?)\*{3}', r'\1', text)
    text = re.sub(r'\*{2}(.+?)\*{2}', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'~~(.+?)~~', r'\1', text)
    text = text.replace('<br>', ' ')
    text = text.replace('\\|', '|')
    return text


def _normalize(text):
    """비교용 정규화 — 공백/줄바꿈 차이 + 마크다운 라운드트립 아티팩트 무시"""
    text = re.sub(r'\s+', ' ', text).strip()
    # * 각주 마커는 마크다운 라운드트립에서 소실되므로 비교 시 무시
    text = text.replace('*', '')
    return text


def _xml_escape(text):
    """XML 텍스트 노드용 이스케이프"""
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    return text


# ============================================================
# 원본 XML 문자열에 직접 텍스트 치환
# ============================================================

def _compute_text_diffs(old_text, new_text):
    """두 텍스트 간 구체적 변경 조각 계산 (replace 연산만).

    Returns:
        list of (old_fragment, new_fragment) tuples
    """
    changes = []
    sm = difflib.SequenceMatcher(None, old_text, new_text)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'replace':
            changes.append((old_text[i1:i2], new_text[j1:j2]))
    return changes


def apply_cell_replacements(raw_xml, replacements, close_tag='</hp:t>'):
    """원본 XML 문자열에서 테이블 셀 텍스트를 직접 치환.

    두 가지 전략을 순차적으로 시도:
      1. 전체 셀 텍스트 매칭 (단일 텍스트 태그 셀)
      2. 프래그먼트 레벨 diff (멀티런 셀 — 변경된 단어만 교체)

    Args:
        raw_xml: 원본 section XML 문자열
        replacements: [(old_text, new_text), ...] — XML 이스케이프된 텍스트
        close_tag: 텍스트 태그 닫기 패턴 (예: '</hp:t>', '</p:t>')

    Returns:
        (modified_xml, applied_count)
    """
    applied = 0
    for old_text, new_text in replacements:
        if not old_text or old_text == new_text:
            continue

        # 전략 1: 전체 텍스트 매칭 (단일 run/t 태그 셀)
        old_pattern = f'>{old_text}{close_tag}'
        new_pattern = f'>{new_text}{close_tag}'

        if old_pattern in raw_xml:
            raw_xml = raw_xml.replace(old_pattern, new_pattern, 1)
            applied += 1
            continue

        # 전략 2: 프래그먼트 레벨 diff (멀티런 셀)
        # 변경된 단어/구절만 찾아서 개별 교체
        changes = _compute_text_diffs(old_text, new_text)
        sub_applied = 0
        for old_frag, new_frag in changes:
            if not old_frag:
                continue
            if old_frag in raw_xml:
                raw_xml = raw_xml.replace(old_frag, new_frag, 1)
                sub_applied += 1
        if sub_applied > 0:
            applied += 1

    return raw_xml, applied


# ============================================================
# 메인 함수
# ============================================================

def _find_section_files(z):
    """ZIP 내부의 모든 Contents/section*.xml 파일을 찾아 숫자순 정렬.

    Args:
        z: zipfile.ZipFile 객체

    Returns:
        list of (section_num, filename) — 숫자순 정렬된 튜플 리스트
    """
    pattern = re.compile(r'^Contents/section(\d+)\.xml$')
    section_files = []
    for name in z.namelist():
        m = pattern.match(name)
        if m:
            section_files.append((int(m.group(1)), name))
    section_files.sort(key=lambda x: x[0])
    return section_files


def smart_replace(original_hwpx, edited_md, output_hwpx=None):
    """원본 HWPX 구조를 보존하며 편집된 마크다운의 테이블 텍스트만 반영.

    다중 섹션(section0.xml, section1.xml, ...)을 모두 처리합니다.
    원본 XML 바이트를 직접 조작하여 lxml 직렬화를 우회합니다.
    """
    if output_hwpx is None:
        base = os.path.splitext(edited_md)[0]
        output_hwpx = base + '_smart.hwpx'

    print(f"스마트 교체 시작:")
    print(f"  원본 HWPX: {original_hwpx}")
    print(f"  편집된 MD: {edited_md}")
    print(f"  출력 HWPX: {output_hwpx}")

    # 1. 마크다운에서 테이블만 추출
    with open(edited_md, 'r', encoding='utf-8') as f:
        md_text = f.read()
    md_tables = parse_markdown_tables(md_text)
    print(f"  마크다운 테이블: {len(md_tables)}개")

    # 2. 원본 HWPX에서 모든 section*.xml 찾기 (숫자순 정렬)
    with open(original_hwpx, 'rb') as f:
        hwpx_bytes = f.read()

    z_in = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')

    section_files = _find_section_files(z_in)
    if not section_files:
        print("오류: Contents/section*.xml을 찾을 수 없습니다.", file=sys.stderr)
        z_in.close()
        sys.exit(1)

    if len(section_files) > 1:
        print(f"  섹션 파일: {len(section_files)}개 ({', '.join(f for _, f in section_files)})")

    # 3. 각 섹션 읽기 및 테이블 추출 (네임스페이스는 첫 섹션에서 감지)
    global NS
    close_tag = '</hp:t>'  # 기본값 — 첫 섹션에서 감지하여 교체

    # 섹션별 데이터: {filename: {'raw_xml': str, 'xml_tables': list, 'table_offset': int}}
    section_data = {}
    all_xml_tables = []  # 전체 테이블 (섹션 순서대로 이어붙임)
    table_to_section = []  # 각 테이블이 속한 섹션 파일명

    for idx, (sec_num, sec_filename) in enumerate(section_files):
        sec_xml_bytes = z_in.read(sec_filename)
        raw_xml = sec_xml_bytes.decode('utf-8')

        # 첫 번째 섹션에서 네임스페이스 + 닫기 태그 감지
        if idx == 0:
            ns_ver = detect_namespace_version(sec_xml_bytes)
            NS = NS_2024.copy() if ns_ver == '2024' else NS_2011.copy()
            if ns_ver == '2024':
                print(f"  네임스페이스: OWPML 2024 감지")
            close_tag = detect_close_tag(raw_xml)

        # lxml으로 분석만 수행 (직렬화 안 함)
        section_root = etree.fromstring(sec_xml_bytes)
        xml_tables = extract_xml_tables(section_root)

        table_offset = len(all_xml_tables)
        section_data[sec_filename] = {
            'raw_xml': raw_xml,
            'xml_tables': xml_tables,
            'table_offset': table_offset,
        }

        for xt in xml_tables:
            all_xml_tables.append(xt)
            table_to_section.append(sec_filename)

    print(f"  XML 테이블: {len(all_xml_tables)}개")

    # 4. 테이블 매칭 및 섹션별 교체 목록 생성
    # per_section_replacements: {filename: [(old_escaped, new_escaped), ...]}
    per_section_replacements = {f: [] for _, f in section_files}
    matched = 0
    skipped = 0

    min_count = min(len(all_xml_tables), len(md_tables))
    for i in range(min_count):
        xt = all_xml_tables[i]
        mt = md_tables[i]

        # 타입 확인 (table↔table, quote↔quote)
        if xt['type'] == 'table' and mt['type'] != 'table':
            skipped += 1
            continue
        if xt['type'] == 'quote' and mt['type'] not in ('quote', 'table'):
            skipped += 1
            continue

        matched += 1
        sec_filename = table_to_section[i]

        # 각 셀 비교
        for row_idx in range(min(xt['row_cnt'], len(mt['cells']))):
            for col_idx in range(min(xt['col_cnt'], len(mt['cells'][row_idx]))):
                old_text = xt['cells'][row_idx][col_idx] if row_idx < len(xt['cells']) else ''
                new_text = _strip_md_format(mt['cells'][row_idx][col_idx])

                if not old_text and not new_text:
                    continue

                # 정규화 비교 — 실제 내용이 다를 때만 교체
                if _normalize(old_text) != _normalize(new_text):
                    # XML 이스케이프
                    old_escaped = _xml_escape(old_text)
                    new_escaped = _xml_escape(new_text)
                    per_section_replacements[sec_filename].append((old_escaped, new_escaped))

    print(f"  테이블 매칭: {matched}개, 건너뜀: {skipped}개")
    total_replacements = sum(len(v) for v in per_section_replacements.values())
    print(f"  교체 대상 셀: {total_replacements}개")

    # 5. 섹션별 원본 XML 문자열에 직접 치환
    # modified_sections: {filename: modified_bytes} — 변경된 섹션만 포함
    modified_sections = {}
    total_applied = 0

    for _, sec_filename in section_files:
        replacements = per_section_replacements[sec_filename]
        if not replacements:
            continue

        raw_xml = section_data[sec_filename]['raw_xml']
        raw_xml, applied = apply_cell_replacements(raw_xml, replacements, close_tag)
        total_applied += applied
        modified_sections[sec_filename] = raw_xml.encode('utf-8')

        if len(section_files) > 1:
            print(f"    {sec_filename}: {applied}개 적용")

    if total_applied > 0:
        print(f"  실제 적용: {total_applied}개")
    else:
        print(f"  변경 사항 없음 — 원본 그대로 복사")

    # 6. HWPX ZIP 재구성 (원본 파일 그대로 + 변경된 섹션만 교체)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            if item.filename in modified_sections:
                z_out.writestr(item.filename, modified_sections[item.filename])
            elif item.filename == 'mimetype':
                z_out.writestr(item, z_in.read(item.filename),
                               compress_type=zipfile.ZIP_STORED)
            else:
                z_out.writestr(item, z_in.read(item.filename))
    z_in.close()

    with open(output_hwpx, 'wb') as f:
        f.write(buf.getvalue())

    print(f"스마트 교체 완료: {output_hwpx}")
    return output_hwpx


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='원본 HWPX 구조 보존 + 마크다운 텍스트 반영 (스마트 교체)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python smart_replace.py 원본.hwpx 편집된.md
  python smart_replace.py 원본.hwpx 편집된.md -o 최종본.hwpx
        """
    )
    parser.add_argument('original', help='원본 HWPX 파일 경로')
    parser.add_argument('markdown', help='편집된 마크다운 파일 경로')
    parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    args = parser.parse_args()

    smart_replace(args.original, args.markdown, args.output)


if __name__ == '__main__':
    main()
