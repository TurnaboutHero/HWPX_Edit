"""
convert.py - HWPX ↔ Markdown 통합 변환 CLI
사용법:
    python convert.py to-md     input.hwpx [-o output.md]
    python convert.py to-hwpx   input.md  [-o output.hwpx] [-r reference.hwpx]
    python convert.py smart     원본.hwpx 편집된.md [-o output.hwpx]
"""
import os
import sys
import argparse
from hwpx_to_md import convert_hwpx_to_md
from md_to_hwpx import convert_md_to_hwpx
from smart_replace import smart_replace


def main():
    parser = argparse.ArgumentParser(
        description='HWPX ↔ Markdown 변환 파이프라인',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  hwpx → markdown:
    python convert.py to-md  신청서.hwpx
    python convert.py to-md  신청서.hwpx -o output/신청서.md

  markdown → hwpx (pypandoc-hwpx 경유):
    python convert.py to-hwpx 사업계획서.md
    python convert.py to-hwpx 사업계획서.md -r 원본양식.hwpx -o 최종본.hwpx

  스마트 교체 (원본 구조 보존, 텍스트만 반영 — 권장):
    python convert.py smart 원본.hwpx 편집된.md
    python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx

  왕복 변환 워크플로:
    python convert.py to-md  원본.hwpx -o 작업폴더/문서.md
    # ... AI로 마크다운 편집 ...
    python convert.py smart 원본.hwpx 작업폴더/문서.md -o 최종본.hwpx
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='변환 방향')

    # to-md 서브커맨드
    md_parser = subparsers.add_parser('to-md', help='HWPX → Markdown 변환')
    md_parser.add_argument('input', help='입력 HWPX 파일')
    md_parser.add_argument('-o', '--output', help='출력 마크다운 파일 경로')
    md_parser.add_argument('--no-images', action='store_true', help='이미지 추출 안 함')

    # to-hwpx 서브커맨드
    hwpx_parser = subparsers.add_parser('to-hwpx', help='Markdown → HWPX 변환 (pypandoc-hwpx)')
    hwpx_parser.add_argument('input', help='입력 마크다운 파일')
    hwpx_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    hwpx_parser.add_argument('-r', '--reference-doc', help='양식 템플릿 원본 HWPX')

    # smart 서브커맨드
    smart_parser = subparsers.add_parser(
        'smart',
        help='스마트 교체 — 원본 HWPX 구조 보존, 텍스트만 반영 (권장)')
    smart_parser.add_argument('original', help='원본 HWPX 파일 경로')
    smart_parser.add_argument('markdown', help='편집된 마크다운 파일 경로')
    smart_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')

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


if __name__ == '__main__':
    main()
