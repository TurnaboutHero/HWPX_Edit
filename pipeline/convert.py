"""
convert.py - HWPX ↔ Markdown 통합 변환 CLI
사용법:
    python convert.py to-md  input.hwpx [-o output.md]
    python convert.py to-hwpx input.md  [-o output.hwpx] [-r reference.hwpx]
"""
import os
import sys
import argparse
from hwpx_to_md import convert_hwpx_to_md
from md_to_hwpx import convert_md_to_hwpx


def main():
    parser = argparse.ArgumentParser(
        description='HWPX ↔ Markdown 변환 파이프라인',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  hwpx → markdown:
    python convert.py to-md  신청서.hwpx
    python convert.py to-md  신청서.hwpx -o output/신청서.md

  markdown → hwpx:
    python convert.py to-hwpx 사업계획서.md
    python convert.py to-hwpx 사업계획서.md -r 원본양식.hwpx -o 최종본.hwpx

  왕복 변환 (양식 보존):
    python convert.py to-md  원본.hwpx -o 작업폴더/문서.md
    # ... 마크다운 편집 ...
    python convert.py to-hwpx 작업폴더/문서.md -r 원본.hwpx -o 최종본.hwpx
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='변환 방향')

    # to-md 서브커맨드
    md_parser = subparsers.add_parser('to-md', help='HWPX → Markdown 변환')
    md_parser.add_argument('input', help='입력 HWPX 파일')
    md_parser.add_argument('-o', '--output', help='출력 마크다운 파일 경로')
    md_parser.add_argument('--no-images', action='store_true', help='이미지 추출 안 함')

    # to-hwpx 서브커맨드
    hwpx_parser = subparsers.add_parser('to-hwpx', help='Markdown → HWPX 변환')
    hwpx_parser.add_argument('input', help='입력 마크다운 파일')
    hwpx_parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    hwpx_parser.add_argument('-r', '--reference-doc', help='양식 템플릿 원본 HWPX')

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == 'to-md':
        convert_hwpx_to_md(args.input, args.output, extract_images=not args.no_images)
    elif args.command == 'to-hwpx':
        convert_md_to_hwpx(args.input, args.output, args.reference_doc)


if __name__ == '__main__':
    main()
