"""
convert.py - HWPX ↔ Markdown 통합 변환 CLI
사용법:
    python convert.py to-md     input.hwpx [-o output.md]
    python convert.py to-hwpx   input.md  [-o output.hwpx] [-r reference.hwpx]
    python convert.py smart     원본.hwpx 편집된.md [-o output.hwpx]
    python convert.py auto      원본.hwpx 편집된.md [-o output.hwpx] [--strip-lineseg]
"""
import os
import sys
import argparse
import io
import zipfile
from lxml import etree
from hwpx_to_md import convert_hwpx_to_md
from md_to_hwpx import convert_md_to_hwpx
from smart_replace import (
    smart_replace,
    parse_markdown_tables,
    parse_markdown_paragraphs,
    extract_xml_tables,
    extract_xml_paragraphs,
    detect_namespace_version,
    _find_section_files,
    NS_2011,
    NS_2024,
)


def auto_detect_and_process(original_hwpx, edited_md, output_hwpx=None, strip_lineseg=False):
    """원본 HWPX와 편집된 마크다운을 비교하여 변경 유형 감지 및 자동 처리.

    변경 유형:
    - 텍스트만 변경: smart_replace 자동 실행
    - 구조 변경 (테이블 개수, 행/열 수): 경고 + smart_replace (텍스트만 반영)

    Args:
        original_hwpx: 원본 HWPX 파일 경로
        edited_md: 편집된 마크다운 파일 경로
        output_hwpx: 출력 HWPX 파일 경로 (None이면 자동 생성)
        strip_lineseg: linesegarray 제거 여부
    """
    if output_hwpx is None:
        base = os.path.splitext(edited_md)[0]
        output_hwpx = base + '_auto.hwpx'

    print(f"자동 변경 감지 시작:")
    print(f"  원본 HWPX: {original_hwpx}")
    print(f"  편집된 MD: {edited_md}")
    print(f"  출력 HWPX: {output_hwpx}")
    print()

    # 1. 마크다운에서 테이블 + 문단 추출
    with open(edited_md, 'r', encoding='utf-8') as f:
        md_text = f.read()
    md_tables = parse_markdown_tables(md_text)
    md_paragraphs = parse_markdown_paragraphs(md_text)

    # 2. 원본 HWPX에서 테이블 + 문단 추출
    with open(original_hwpx, 'rb') as f:
        hwpx_bytes = f.read()

    z_in = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')
    section_files = _find_section_files(z_in)

    if not section_files:
        print("오류: Contents/section*.xml을 찾을 수 없습니다.", file=sys.stderr)
        z_in.close()
        sys.exit(1)

    # 네임스페이스 감지
    _, first_section = section_files[0]
    sec_xml_bytes = z_in.read(first_section)
    ns_ver = detect_namespace_version(sec_xml_bytes)
    NS = NS_2024.copy() if ns_ver == '2024' else NS_2011.copy()

    # 모든 섹션에서 테이블 + 문단 추출
    all_xml_tables = []
    all_xml_paragraphs = []

    for _, sec_filename in section_files:
        sec_xml_bytes = z_in.read(sec_filename)
        section_root = etree.fromstring(sec_xml_bytes)
        xml_tables = extract_xml_tables(section_root)
        xml_paragraphs = extract_xml_paragraphs(section_root)
        all_xml_tables.extend(xml_tables)
        all_xml_paragraphs.extend(xml_paragraphs)

    z_in.close()

    # 3. 변경 유형 감지
    print("변경 사항 분석:")
    print(f"  테이블: 원본 {len(all_xml_tables)}개 → 편집 {len(md_tables)}개")
    print(f"  문단: 원본 {len(all_xml_paragraphs)}개 → 편집 {len(md_paragraphs)}개")
    print()

    has_structural_changes = False
    warnings = []

    # 테이블 개수 변화
    if len(all_xml_tables) != len(md_tables):
        has_structural_changes = True
        diff = len(md_tables) - len(all_xml_tables)
        if diff > 0:
            warnings.append(f"테이블 {diff}개 추가됨")
        else:
            warnings.append(f"테이블 {-diff}개 삭제됨")

    # 테이블 행/열 수 변화
    for i in range(min(len(all_xml_tables), len(md_tables))):
        xml_tbl = all_xml_tables[i]
        md_tbl = md_tables[i]

        # 인용문(1×1 테이블)은 구조 변경 감지 제외
        if xml_tbl['type'] == 'quote':
            continue

        xml_rows = xml_tbl['row_cnt']
        xml_cols = xml_tbl['col_cnt']
        md_rows = len(md_tbl['cells'])
        md_cols = len(md_tbl['cells'][0]) if md_tbl['cells'] else 0

        if xml_rows != md_rows or xml_cols != md_cols:
            has_structural_changes = True
            warnings.append(f"테이블 #{i+1}: {xml_rows}×{xml_cols} → {md_rows}×{md_cols} (구조 변경)")

    # 문단 개수 변화
    if len(all_xml_paragraphs) != len(md_paragraphs):
        para_diff = len(md_paragraphs) - len(all_xml_paragraphs)
        if abs(para_diff) > 2:  # 미세한 차이는 무시 (파싱 휴리스틱 차이)
            has_structural_changes = True
            if para_diff > 0:
                warnings.append(f"문단 {para_diff}개 추가됨")
            else:
                warnings.append(f"문단 {-para_diff}개 삭제됨")

    # 4. 결과 출력 및 처리 경로 선택
    if has_structural_changes:
        print("[경고] 구조 변경 감지:")
        for warning in warnings:
            print(f"    - {warning}")
        print()
        print("스마트 교체 모드로 진행합니다.")
        print("주의: 구조 변경 사항(행/열 추가/삭제)은 반영되지 않으며,")
        print("      텍스트 변경만 원본 HWPX 구조에 반영됩니다.")
        print()
    else:
        print("[OK] 텍스트만 변경됨 (구조 변경 없음)")
        print("스마트 교체 모드로 진행합니다.")
        print()

    # 5. smart_replace 실행
    result_path = smart_replace(original_hwpx, edited_md, output_hwpx)

    # 6. linesegarray 제거 (옵션)
    if strip_lineseg:
        print()
        print("linesegarray 제거 중...")
        _strip_linesegarray(result_path)
        print(f"linesegarray 제거 완료: {result_path}")

    return result_path


def _strip_linesegarray(hwpx_path):
    """HWPX 파일에서 모든 linesegarray 태그 제거.

    linesegarray는 줄 나눔 정보를 담고 있으나, 텍스트 변경 시 무효화되어
    한글에서 렌더링 오류를 유발할 수 있습니다. 제거 시 한글이 자동으로 재계산합니다.
    """
    with open(hwpx_path, 'rb') as f:
        hwpx_bytes = f.read()

    z_in = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')
    section_files = _find_section_files(z_in)

    modified_sections = {}

    for _, sec_filename in section_files:
        raw_xml = z_in.read(sec_filename).decode('utf-8')

        # linesegarray 태그 제거 (내용 포함 형태 + 자기 닫힘 형태 모두)
        import re
        modified_xml = re.sub(r'<[\w]+:linesegarray[^>]*>.*?</[\w]+:linesegarray>', '', raw_xml, flags=re.DOTALL)
        modified_xml = re.sub(r'<[\w]+:linesegarray[^>]*?/>', '', modified_xml)

        if modified_xml != raw_xml:
            modified_sections[sec_filename] = modified_xml.encode('utf-8')

    # 수정된 섹션으로 HWPX 재구성
    if modified_sections:
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

        with open(hwpx_path, 'wb') as f:
            f.write(buf.getvalue())
    else:
        z_in.close()


def main():
    parser = argparse.ArgumentParser(
        description='HWPX ↔ Markdown 변환 파이프라인',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  hwpx -> markdown:
    python convert.py to-md  신청서.hwpx
    python convert.py to-md  신청서.hwpx -o output/신청서.md

  markdown -> hwpx (pypandoc-hwpx 경유):
    python convert.py to-hwpx 사업계획서.md
    python convert.py to-hwpx 사업계획서.md -r 원본양식.hwpx -o 최종본.hwpx

  스마트 교체 (원본 구조 보존, 텍스트만 반영 - 권장):
    python convert.py smart 원본.hwpx 편집된.md
    python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx

  자동 변경 감지 및 처리:
    python convert.py auto 원본.hwpx 편집된.md
    python convert.py auto 원본.hwpx 편집된.md -o 최종본.hwpx --strip-lineseg

  왕복 변환 워크플로:
    python convert.py to-md  원본.hwpx -o 작업폴더/문서.md
    # ... AI로 마크다운 편집 ...
    python convert.py smart 원본.hwpx 작업폴더/문서.md -o 최종본.hwpx
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='변환 방향')

    # to-md 서브커맨드
    md_parser = subparsers.add_parser('to-md', help='HWPX -> Markdown 변환')
    md_parser.add_argument('input', help='입력 HWPX 파일')
    md_parser.add_argument('-o', '--output', help='출력 마크다운 파일 경로')
    md_parser.add_argument('--no-images', action='store_true', help='이미지 추출 안 함')

    # to-hwpx 서브커맨드
    hwpx_parser = subparsers.add_parser('to-hwpx', help='Markdown -> HWPX 변환 (pypandoc-hwpx)')
    hwpx_parser.add_argument('input', help='입력 마크다운 파일')
    hwpx_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    hwpx_parser.add_argument('-r', '--reference-doc', help='양식 템플릿 원본 HWPX')

    # smart 서브커맨드
    smart_parser = subparsers.add_parser(
        'smart',
        help='스마트 교체 - 원본 HWPX 구조 보존, 텍스트만 반영 (권장)')
    smart_parser.add_argument('original', help='원본 HWPX 파일 경로')
    smart_parser.add_argument('markdown', help='편집된 마크다운 파일 경로')
    smart_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')

    # auto 서브커맨드
    auto_parser = subparsers.add_parser(
        'auto',
        help='자동 변경 감지 - 변경 유형 분석 후 적절한 처리 경로 선택')
    auto_parser.add_argument('original', help='원본 HWPX 파일 경로')
    auto_parser.add_argument('markdown', help='편집된 마크다운 파일 경로')
    auto_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    auto_parser.add_argument('--strip-lineseg', action='store_true',
                             help='linesegarray 제거 (기본: 유지)')

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == 'to-md':
        convert_hwpx_to_md(args.input, args.output, extract_images=not args.no_images)
    elif args.command == 'to-hwpx':
        convert_md_to_hwpx(args.input, args.output, args.reference_doc)
    elif args.command == 'smart':
        smart_replace(args.original, args.markdown, args.output)
    elif args.command == 'auto':
        auto_detect_and_process(args.original, args.markdown, args.output,
                                strip_lineseg=args.strip_lineseg)


if __name__ == '__main__':
    main()
