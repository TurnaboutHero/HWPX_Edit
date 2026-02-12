"""
md_to_hwpx.py - Markdown to HWPX 변환 래퍼
pypandoc-hwpx를 활용하여 마크다운을 HWPX로 변환합니다.
원본 hwpx를 --reference-doc으로 사용하여 양식을 보존합니다.
"""
import os
import sys
import argparse
import subprocess


def convert_md_to_hwpx(input_path, output_path=None, reference_doc=None):
    """마크다운을 hwpx로 변환"""
    if output_path is None:
        base = os.path.splitext(input_path)[0]
        output_path = base + '.hwpx'

    cmd = ['pypandoc-hwpx', input_path, '-o', output_path]

    if reference_doc:
        cmd.extend(['--reference-doc', reference_doc])

    print(f"변환 중: {input_path} → {output_path}")
    if reference_doc:
        print(f"양식 템플릿: {reference_doc}")

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        print(f"변환 실패 (exit code: {result.returncode})", file=sys.stderr)
        sys.exit(1)

    print(f"변환 완료: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description='Markdown → HWPX 변환기 (pypandoc-hwpx 래퍼)')
    parser.add_argument('input', help='입력 마크다운 파일 경로')
    parser.add_argument('-o', '--output', help='출력 HWPX 파일 경로')
    parser.add_argument('-r', '--reference-doc', help='양식 템플릿으로 사용할 원본 HWPX 파일')
    args = parser.parse_args()

    convert_md_to_hwpx(args.input, args.output, args.reference_doc)


if __name__ == '__main__':
    main()
