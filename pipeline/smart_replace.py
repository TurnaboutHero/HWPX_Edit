"""
smart_replace.py - 원본 HWPX 구조 보존 + 마크다운 텍스트 반영

원본 HWPX의 바이트를 그대로 보존하면서, 편집된 마크다운에서 변경된
테이블 셀 텍스트 및 일반 문단 텍스트를 원본 XML 문자열에 직접 치환합니다.

lxml은 분석(파싱)에만 사용하고, 직렬화는 하지 않습니다.
이를 통해 CRLF, 속성 순서, 네임스페이스 등이 원본과 100% 동일하게 보존됩니다.

주요 장점:
  - 표 서식 완벽 보존 (셀 병합, 너비, 높이, 테두리, 배경색)
  - 일반 문단 텍스트 교체 지원
  - 원본 XML 바이트 수준 보존 (lxml 직렬화 우회)
  - pypandoc-hwpx 버그 우회

제한사항:
  - 구조 변경(행/열 추가/삭제)은 반영 불가
  - 제목, 이미지 참조, 인용문 내 텍스트는 교체 대상에서 제외
"""
import os
import re
import sys
import argparse
import zipfile
import io
import difflib
import tempfile
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


def parse_markdown_paragraphs(md_text):
    """마크다운에서 일반 텍스트 문단을 순서대로 추출.

    ``hwpx_to_md.py``는 HWPX 문단 하나를 Markdown 한 줄로 내보냅니다.
    smart_replace는 원본 XML 문단과 1:1로 맞춰야 하므로, 연속된 줄을
    임의로 합치지 않고 비어 있지 않은 일반 텍스트 줄 하나를 문단 하나로
    취급합니다. 사용자가 줄을 수동으로 감싸서 문단 수가 달라지면
    smart_replace 단계에서 중단되어 잘못된 위치에 텍스트가 들어가지 않게
    합니다.

    Returns:
        list of str: 문단 텍스트 목록
    """
    paragraphs = []
    lines = md_text.split('\n')
    i = 0
    in_table = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 빈 줄 — 현재 문단 종료
        if not stripped:
            in_table = False
            i += 1
            continue

        # 테이블 행 (|로 시작하거나 구분선 |---|)
        if stripped.startswith('|') or (in_table and '|' in stripped):
            in_table = True
            i += 1
            continue

        # 테이블 시작 감지 (다음 줄이 구분선)
        if '|' in stripped and i + 1 < len(lines):
            next_stripped = lines[i + 1].strip()
            if re.match(r'^\|[\s\-:|]+\|$', next_stripped):
                in_table = True
                i += 1
                continue

        in_table = False

        # 제목 (#으로 시작)
        if stripped.startswith('#'):
            i += 1
            continue

        # 인용문 (>로 시작)
        if stripped.startswith('>'):
            i += 1
            continue

        # 이미지 (![으로 시작)
        if stripped.startswith('!['):
            i += 1
            continue

        # 구분선 (---, ***, ___)
        if re.match(r'^[-*_]{3,}\s*$', stripped):
            i += 1
            continue

        # HTML 주석 (<!-- -->)
        if stripped.startswith('<!--'):
            i += 1
            continue

        # 수식 블록 ($$ ... $$) — 시작~끝 모두 건너뜀
        if stripped.startswith('$$'):
            i += 1
            # $$ 내부 줄도 건너뛰기
            while i < len(lines) and not lines[i].strip().startswith('$$'):
                i += 1
            if i < len(lines):
                i += 1  # 닫는 $$ 줄도 건너뜀
            continue

        # 각주/미주 정의 ([^N]: ...)
        if re.match(r'^\[\^.+?\]:', stripped):
            i += 1
            continue

        # 양식 개체 패턴 ([x], [ ], (o), ( ), [콤보:, [버튼:, [입력란:)
        if re.match(r'^\[[ x]\]\s', stripped) or re.match(r'^\([ o]\)\s', stripped):
            i += 1
            continue
        if re.match(r'^\[(콤보|버튼|입력란):', stripped):
            i += 1
            continue

        paragraphs.append(stripped)
        i += 1

    return paragraphs


def parse_protected_markdown_lines(md_text):
    """smart_replace가 직접 반영하지 않는 Markdown 라인을 추출.

    이미지, 제목, 글상자/글맵시 인용 라인, OLE/메타 주석, 수식 블록,
    양식 개체 등은 현재 텍스트 치환 대상이 아닙니다. 이런 줄이 바뀌면
    조용히 무시하지 않고 중단하기 위해 원본/편집본 비교에 사용합니다.
    """
    protected = []
    lines = md_text.split('\n')
    i = 0
    in_table = False
    in_equation = False

    while i < len(lines):
        stripped = lines[i].strip()

        if not stripped:
            in_table = False
            i += 1
            continue

        if in_equation:
            protected.append(stripped)
            if stripped.startswith('$$'):
                in_equation = False
            i += 1
            continue

        if stripped.startswith('$$'):
            protected.append(stripped)
            if stripped.count('$$') == 1:
                in_equation = True
            i += 1
            continue

        if stripped.startswith('|') or (in_table and '|' in stripped):
            in_table = True
            i += 1
            continue

        if '|' in stripped and i + 1 < len(lines):
            next_stripped = lines[i + 1].strip()
            if re.match(r'^\|[\s\-:|]+\|$', next_stripped):
                in_table = True
                i += 1
                continue

        in_table = False

        if (
            stripped.startswith('#') or
            stripped.startswith('>') or
            stripped.startswith('![') or
            stripped.startswith('<!--') or
            re.match(r'^[-*_]{3,}\s*$', stripped) or
            re.match(r'^\[\^.+?\]:', stripped) or
            re.match(r'^\[[ x]\]\s', stripped) or
            re.match(r'^\([ o]\)\s', stripped) or
            re.match(r'^\[(콤보|버튼|입력란):', stripped)
        ):
            protected.append(stripped)

        i += 1

    return protected


def _normalize_protected_markdown_line(line):
    """비교용 보호 라인 정규화."""
    image_match = re.match(r'^!\[([^\]]*)\]\([^)]+\)$', line)
    if image_match:
        return f"![{image_match.group(1)}]"
    return line


def collect_protected_markup_changes(original_md, edited_md):
    """지원하지 않는 Markdown 라인 변경 사항을 찾습니다."""
    original_lines = [
        _normalize_protected_markdown_line(line)
        for line in parse_protected_markdown_lines(original_md)
    ]
    edited_lines = [
        _normalize_protected_markdown_line(line)
        for line in parse_protected_markdown_lines(edited_md)
    ]

    if original_lines == edited_lines:
        return []

    changes = []
    max_len = max(len(original_lines), len(edited_lines))
    for idx in range(max_len):
        old = original_lines[idx] if idx < len(original_lines) else ''
        new = edited_lines[idx] if idx < len(edited_lines) else ''
        if old != new:
            changes.append({
                'index': idx + 1,
                'old': old,
                'new': new,
            })
    return changes


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


def _is_heading_para(para):
    """hp:p가 제목(heading) 문단인지 휴리스틱으로 판별.

    hwpx_to_md.py는 제목을 # 마크다운으로 변환하고,
    parse_markdown_paragraphs는 #으로 시작하는 줄을 건너뛰므로
    XML 쪽에서도 제목 문단을 건너뛰어야 정렬이 맞습니다.

    판별 기준:
      1. outlineLevel 속성이 있는 문단 (OWPML 표준 제목)
      2. paraStyleIDRef가 '개요' 또는 'Heading'을 포함하는 문단
    """
    # 방법 1: paraPr의 outlineLevel 확인
    for pr in para.findall('hp:paraPr', NS):
        if pr.get('outlineLevel'):
            return True

    # 방법 2: styleIDRef 패턴 확인 (개요 1~9, Heading 1~9)
    style_ref = para.get('styleIDRef', '')
    if style_ref and re.match(r'(개요|Heading|heading)\s*\d', style_ref):
        return True

    return False


def extract_xml_paragraphs(section_root):
    """section XML에서 테이블/이미지/제목을 제외한 최상위 문단 텍스트 추출.

    hwpx_to_md.py의 _process_section()과 동일한 순서로 순회하여
    마크다운 문단과 1:1 매칭할 수 있도록 합니다.

    Returns:
        list of str: 비어있지 않은 순수 텍스트 문단 목록
    """
    paragraphs = []
    for para in section_root.findall('hp:p', NS):
        # 테이블을 포함한 문단은 건너뜀 (이미 테이블로 처리)
        if para.findall('.//hp:tbl', NS):
            continue
        # 이미지를 포함한 문단은 건너뜀
        if para.findall('.//hp:pic', NS):
            continue
        # 제목 문단은 건너뜀 (마크다운에서 # 으로 변환되어 제외됨)
        if _is_heading_para(para):
            continue
        text = _get_para_text(para)
        if not text.strip():
            continue
        paragraphs.append(text.strip())
    return paragraphs


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


def _display_width(text):
    """대략적인 표시 폭 계산.

    정확한 조판 폭은 글꼴/크기/문단 폭에 따라 달라지지만, CJK 문자는
    영문보다 넓게 잡아 긴 답변 리스크를 보수적으로 감지합니다.
    """
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 127 else 1
    return width


def collect_text_growth_warnings(old_items, new_items, kind='문단',
                                 ratio_threshold=1.6, delta_threshold=60,
                                 width_threshold=180):
    """텍스트 길이 증가로 레이아웃이 밀릴 가능성이 큰 항목을 찾습니다.

    Returns:
        list of dict: index, kind, old_width, new_width, delta, ratio, preview
    """
    warnings = []
    for idx, (old_text, new_text) in enumerate(zip(old_items, new_items), 1):
        old_plain = _strip_md_format(old_text)
        new_plain = _strip_md_format(new_text)
        old_width = _display_width(old_plain)
        new_width = _display_width(new_plain)
        delta = new_width - old_width
        ratio = (new_width / old_width) if old_width else float('inf')

        if new_width >= width_threshold and delta >= delta_threshold and ratio >= ratio_threshold:
            warnings.append({
                'index': idx,
                'kind': kind,
                'old_width': old_width,
                'new_width': new_width,
                'delta': delta,
                'ratio': ratio,
                'preview': new_plain[:80],
            })

    return warnings


def format_growth_warning(warning):
    """길이 증가 경고를 CLI/대시보드 공통 메시지로 변환."""
    ratio = warning['ratio']
    ratio_text = '∞' if ratio == float('inf') else f"{ratio:.1f}배"
    return (
        f"{warning['kind']} #{warning['index']}: "
        f"폭 {warning['old_width']} -> {warning['new_width']} "
        f"(+{warning['delta']}, {ratio_text})"
    )


def format_protected_markup_change(change):
    """지원하지 않는 Markdown 변경 경고를 메시지로 변환."""
    old = change['old'] or '(없음)'
    new = change['new'] or '(없음)'
    return f"보호 라인 #{change['index']}: {old[:60]} -> {new[:60]}"


def render_original_markdown_for_checks(hwpx_path):
    """원본 HWPX를 Markdown으로 렌더링해 보호 라인 비교에 사용."""
    from hwpx_to_md import HwpxToMarkdown

    with tempfile.TemporaryDirectory(prefix='hwpx_check_') as tmp_dir:
        converter = HwpxToMarkdown(
            hwpx_path,
            output_dir=tmp_dir,
            extract_images=False,
        )
        return converter.convert()


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


def _replace_in_text_node(raw_xml, old_frag, new_frag):
    """텍스트 노드(> ... <) 내부에서만 프래그먼트를 교체.

    XML 속성값이나 태그 이름이 아닌, 실제 텍스트 콘텐츠 안에서만
    교체가 일어나도록 보장합니다.

    Returns:
        (modified_xml, success: bool)
    """
    start = 0
    while True:
        idx = raw_xml.find(old_frag, start)
        if idx == -1:
            return raw_xml, False

        # 이 위치가 텍스트 노드 내부인지 확인
        # 조건: idx 앞의 마지막 '>'와 idx 사이에 '<'가 없어야 함
        last_gt = raw_xml.rfind('>', 0, idx)
        if last_gt == -1:
            start = idx + 1
            continue

        between = raw_xml[last_gt + 1:idx]
        if '<' in between:
            start = idx + 1
            continue

        # 프래그먼트가 태그 경계를 넘지 않는지 확인
        frag_end = idx + len(old_frag)
        next_lt = raw_xml.find('<', idx)
        if next_lt != -1 and next_lt < frag_end:
            start = idx + 1
            continue

        # 안전 — 텍스트 노드 내부에서 교체
        raw_xml = raw_xml[:idx] + new_frag + raw_xml[frag_end:]
        return raw_xml, True


def apply_cell_replacements(raw_xml, replacements, close_tag='</hp:t>'):
    """원본 XML 문자열에서 테이블 셀 텍스트를 직접 치환.

    두 가지 전략을 순차적으로 시도:
      1. 전체 셀 텍스트 매칭 (단일 텍스트 태그 셀)
      2. 프래그먼트 레벨 diff (멀티런 셀 — 텍스트 노드 안에서만 교체)

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
        # 텍스트 노드 내부에서만 교체 (XML 속성/태그 보호)
        changes = _compute_text_diffs(old_text, new_text)
        sub_applied = 0
        for old_frag, new_frag in changes:
            if not old_frag or len(old_frag) < 2:
                continue
            raw_xml, ok = _replace_in_text_node(raw_xml, old_frag, new_frag)
            if ok:
                sub_applied += 1
        if sub_applied > 0:
            applied += 1

    return raw_xml, applied


def apply_para_replacements(raw_xml, replacements, close_tag='</hp:t>'):
    """원본 XML 문자열에서 문단 텍스트를 직접 치환.

    테이블 셀과 달리 프래그먼트 diff를 사용하지 않음.
    문단 전체 텍스트가 하나의 텍스트 노드 안에 있을 때만 교체하여
    XML 구조 파손을 방지합니다. 텍스트 뒤에 ``hp:tab`` 같은 자식 태그가
    붙어 있어도 텍스트 노드 내부의 실제 문장만 교체합니다.

    Args:
        raw_xml: 원본 section XML 문자열
        replacements: [(old_text, new_text), ...] — XML 이스케이프된 텍스트
        close_tag: 텍스트 태그 닫기 패턴

    Returns:
        (modified_xml, applied_count)
    """
    applied = 0
    for old_text, new_text in replacements:
        if not old_text or old_text == new_text:
            continue

        raw_xml, ok = _replace_in_text_node(raw_xml, old_text, new_text)
        if ok:
            applied += 1

    return raw_xml, applied


def strip_linesegarray(hwpx_path):
    """HWPX 파일에서 section XML의 linesegarray 태그를 제거.

    linesegarray는 줄 배치 캐시입니다. 텍스트 길이가 바뀐 뒤에도 남아 있으면
    한글에서 예전 줄 위치를 재사용해 글자가 겹쳐 보일 수 있습니다.
    제거하면 한글이 문서를 열 때 줄 배치를 다시 계산합니다.

    Returns:
        int: 제거된 linesegarray 태그 수
    """
    with open(hwpx_path, 'rb') as f:
        original_bytes = f.read()

    z_in = zipfile.ZipFile(io.BytesIO(original_bytes), 'r')
    removed_count = 0
    modified_sections = {}

    for item in z_in.infolist():
        if not re.match(r'^Contents/section\d+\.xml$', item.filename):
            continue

        raw_xml = z_in.read(item.filename).decode('utf-8')
        modified_xml, count = re.subn(
            r'<[\w]+:linesegarray[^>]*>.*?</[\w]+:linesegarray>',
            '',
            raw_xml,
            flags=re.DOTALL,
        )
        modified_xml, self_closing_count = re.subn(
            r'<[\w]+:linesegarray[^>]*/>',
            '',
            modified_xml,
        )
        count += self_closing_count

        if count:
            removed_count += count
            modified_sections[item.filename] = modified_xml.encode('utf-8')

    if not modified_sections:
        z_in.close()
        return 0

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            if item.filename in modified_sections:
                z_out.writestr(item, modified_sections[item.filename])
            elif item.filename == 'mimetype':
                z_out.writestr(item, z_in.read(item.filename),
                               compress_type=zipfile.ZIP_STORED)
            else:
                z_out.writestr(item, z_in.read(item.filename))
    z_in.close()

    with open(hwpx_path, 'wb') as f:
        f.write(buf.getvalue())

    return removed_count


def validate_hwpx_integrity(hwpx_path, require_no_lineseg=False):
    """HWPX 산출물이 다시 열리고 section XML이 파싱되는지 검증."""
    result = {
        'ok': True,
        'errors': [],
        'section_count': 0,
        'linesegarray_count': 0,
    }

    try:
        with zipfile.ZipFile(hwpx_path, 'r') as z:
            bad_member = z.testzip()
            if bad_member:
                result['errors'].append(f'ZIP 멤버 손상: {bad_member}')
            section_files = _find_section_files(z)
            result['section_count'] = len(section_files)

            if not section_files:
                result['errors'].append('Contents/section*.xml 파일이 없습니다.')

            for _, section_file in section_files:
                raw_xml = z.read(section_file)
                try:
                    etree.fromstring(raw_xml)
                except etree.XMLSyntaxError as exc:
                    result['errors'].append(f'{section_file} XML 파싱 실패: {exc}')

                raw_text = raw_xml.decode('utf-8')
                result['linesegarray_count'] += len(
                    re.findall(r'<[\w]+:linesegarray\b', raw_text)
                )

            if require_no_lineseg and result['linesegarray_count'] > 0:
                result['errors'].append(
                    f"linesegarray가 {result['linesegarray_count']}개 남아 있습니다."
                )
    except Exception as exc:
        result['errors'].append(f'HWPX ZIP 검증 실패: {exc}')

    result['ok'] = not result['errors']
    return result


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


def smart_replace(original_hwpx, edited_md, output_hwpx=None, strip_lineseg=True,
                  allow_layout_risk=False):
    """원본 HWPX 구조를 보존하며 편집된 마크다운의 텍스트를 반영.

    테이블 셀 텍스트와 일반 문단 텍스트를 모두 교체합니다.
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

    # 1. 마크다운에서 테이블 + 문단 추출
    with open(edited_md, 'r', encoding='utf-8') as f:
        md_text = f.read()
    md_tables = parse_markdown_tables(md_text)
    md_paragraphs = parse_markdown_paragraphs(md_text)
    print(f"  마크다운 테이블: {len(md_tables)}개, 문단: {len(md_paragraphs)}개")

    original_md_for_checks = render_original_markdown_for_checks(original_hwpx)
    protected_changes = collect_protected_markup_changes(original_md_for_checks, md_text)
    if protected_changes:
        print("  보호된 Markdown 라인이 변경되었습니다:")
        for change in protected_changes[:5]:
            print(f"    - {format_protected_markup_change(change)}")
        if len(protected_changes) > 5:
            print(f"    - 외 {len(protected_changes) - 5}개")
        raise ValueError(
            "현재 smart_replace는 이미지, 제목, 글상자, OLE, 수식, 양식 개체 "
            "변경을 반영하지 않습니다. 해당 라인을 원본과 같게 유지하세요."
        )

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

    # 섹션별 데이터: {filename: {'raw_xml': str, 'xml_tables': list, 'xml_paragraphs': list, 'table_offset': int, 'para_offset': int}}
    section_data = {}
    all_xml_tables = []  # 전체 테이블 (섹션 순서대로 이어붙임)
    all_xml_paragraphs = []  # 전체 문단 (섹션 순서대로 이어붙임)
    table_to_section = []  # 각 테이블이 속한 섹션 파일명
    para_to_section = []  # 각 문단이 속한 섹션 파일명

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
        xml_paragraphs = extract_xml_paragraphs(section_root)

        table_offset = len(all_xml_tables)
        para_offset = len(all_xml_paragraphs)
        section_data[sec_filename] = {
            'raw_xml': raw_xml,
            'xml_tables': xml_tables,
            'xml_paragraphs': xml_paragraphs,
            'table_offset': table_offset,
            'para_offset': para_offset,
        }

        for xt in xml_tables:
            all_xml_tables.append(xt)
            table_to_section.append(sec_filename)

        for xp in xml_paragraphs:
            all_xml_paragraphs.append(xp)
            para_to_section.append(sec_filename)

    print(f"  XML 테이블: {len(all_xml_tables)}개, 문단: {len(all_xml_paragraphs)}개")

    if len(md_tables) != len(all_xml_tables):
        z_in.close()
        raise ValueError(
            "테이블 수가 원본과 편집본에서 다릅니다. "
            f"원본 XML={len(all_xml_tables)}개, Markdown={len(md_tables)}개. "
            "smart_replace는 테이블 추가/삭제를 지원하지 않습니다."
        )

    if len(md_paragraphs) != len(all_xml_paragraphs):
        z_in.close()
        raise ValueError(
            "문단 수가 원본과 편집본에서 다릅니다. "
            f"원본 XML={len(all_xml_paragraphs)}개, Markdown={len(md_paragraphs)}개. "
            "HWPX 한 문단은 Markdown 한 줄로 유지해야 합니다."
        )

    table_old_items = []
    table_new_items = []
    for xt, mt in zip(all_xml_tables, md_tables):
        for row_idx in range(min(xt['row_cnt'], len(mt['cells']))):
            for col_idx in range(min(xt['col_cnt'], len(mt['cells'][row_idx]))):
                table_old_items.append(xt['cells'][row_idx][col_idx])
                table_new_items.append(mt['cells'][row_idx][col_idx])

    growth_warnings = []
    growth_warnings.extend(
        collect_text_growth_warnings(table_old_items, table_new_items, kind='셀')
    )
    growth_warnings.extend(
        collect_text_growth_warnings(all_xml_paragraphs, md_paragraphs, kind='문단')
    )

    if growth_warnings:
        print("  레이아웃 주의: 텍스트가 크게 길어진 항목이 있습니다.")
        for warning in growth_warnings[:5]:
            print(f"    - {format_growth_warning(warning)}")
        if len(growth_warnings) > 5:
            print(f"    - 외 {len(growth_warnings) - 5}개")
        if not allow_layout_risk:
            z_in.close()
            raise ValueError(
                "텍스트 길이 증가로 레이아웃 위험이 감지되어 생성을 중단했습니다. "
                "내용을 줄이거나, 위험을 확인한 뒤 --allow-layout-risk 옵션을 사용하세요."
            )

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

    # 4.5. 문단 매칭 및 섹션별 교체 목록 생성
    per_section_para_replacements = {f: [] for _, f in section_files}
    para_matched = 0
    para_changed = 0

    min_para_count = min(len(all_xml_paragraphs), len(md_paragraphs))
    for i in range(min_para_count):
        xml_para_text = all_xml_paragraphs[i]
        md_para_text = _strip_md_format(md_paragraphs[i])

        # 정규화 비교 — 실제 내용이 다를 때만 교체
        if _normalize(xml_para_text) == _normalize(md_para_text):
            para_matched += 1
            continue

        para_matched += 1
        para_changed += 1
        sec_filename = para_to_section[i]

        # XML 이스케이프
        old_escaped = _xml_escape(xml_para_text)
        new_escaped = _xml_escape(md_para_text)
        per_section_para_replacements[sec_filename].append((old_escaped, new_escaped))

    total_para_replacements = sum(len(v) for v in per_section_para_replacements.values())
    print(f"  문단 매칭: {para_matched}개, 변경: {para_changed}개")
    if total_para_replacements > 0:
        print(f"  교체 대상 문단: {total_para_replacements}개")

    # 5. 섹션별 원본 XML 문자열에 직접 치환 (테이블 + 문단)
    # modified_sections: {filename: modified_bytes} — 변경된 섹션만 포함
    modified_sections = {}
    total_applied = 0
    total_para_applied = 0

    for _, sec_filename in section_files:
        cell_replacements = per_section_replacements[sec_filename]
        para_replacements = per_section_para_replacements[sec_filename]

        if not cell_replacements and not para_replacements:
            continue

        raw_xml = section_data[sec_filename]['raw_xml']

        # 테이블 셀 교체
        cell_applied = 0
        if cell_replacements:
            raw_xml, cell_applied = apply_cell_replacements(raw_xml, cell_replacements, close_tag)
            total_applied += cell_applied

        # 문단 텍스트 교체 (전체 매칭만 — 프래그먼트 diff 금지)
        para_applied = 0
        if para_replacements:
            raw_xml, para_applied = apply_para_replacements(raw_xml, para_replacements, close_tag)
            total_para_applied += para_applied

        modified_sections[sec_filename] = raw_xml.encode('utf-8')

        if len(section_files) > 1:
            parts = []
            if cell_applied > 0:
                parts.append(f"셀 {cell_applied}개")
            if para_applied > 0:
                parts.append(f"문단 {para_applied}개")
            if parts:
                print(f"    {sec_filename}: {', '.join(parts)} 적용")

    if total_applied > 0 or total_para_applied > 0:
        parts = []
        if total_applied > 0:
            parts.append(f"셀 {total_applied}개")
        if total_para_applied > 0:
            parts.append(f"문단 {total_para_applied}개")
        print(f"  실제 적용: {', '.join(parts)}")
    else:
        print(f"  변경 사항 없음 — 원본 그대로 복사")

    expected_cell_replacements = sum(len(v) for v in per_section_replacements.values())
    expected_para_replacements = sum(len(v) for v in per_section_para_replacements.values())
    if total_applied != expected_cell_replacements:
        z_in.close()
        raise ValueError(
            "일부 테이블 셀 교체를 적용하지 못했습니다. "
            f"예상={expected_cell_replacements}개, 적용={total_applied}개."
        )
    if total_para_applied != expected_para_replacements:
        z_in.close()
        raise ValueError(
            "일부 문단 교체를 적용하지 못했습니다. "
            f"예상={expected_para_replacements}개, 적용={total_para_applied}개."
        )

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

    if strip_lineseg:
        removed = strip_linesegarray(output_hwpx)
        if removed:
            print(f"  linesegarray 제거: {removed}개")

    validation = validate_hwpx_integrity(output_hwpx, require_no_lineseg=strip_lineseg)
    if not validation['ok']:
        raise ValueError(
            "산출물 검증 실패: " + " / ".join(validation['errors'])
        )
    print(
        f"  산출물 검증: 섹션 {validation['section_count']}개, "
        f"linesegarray {validation['linesegarray_count']}개"
    )

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
    parser.add_argument('--keep-lineseg', action='store_true',
                        help='줄 배치 캐시(linesegarray)를 유지합니다')
    parser.add_argument('--allow-layout-risk', action='store_true',
                        help='긴 텍스트로 인한 레이아웃 위험을 확인하고 생성을 허용합니다')
    args = parser.parse_args()

    smart_replace(args.original, args.markdown, args.output,
                  strip_lineseg=not args.keep_lineseg,
                  allow_layout_risk=args.allow_layout_risk)


if __name__ == '__main__':
    main()
