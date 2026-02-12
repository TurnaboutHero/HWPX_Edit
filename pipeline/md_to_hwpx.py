"""
md_to_hwpx.py - Markdown to HWPX 변환 래퍼
pypandoc-hwpx를 활용하여 마크다운을 HWPX로 변환합니다.
원본 hwpx를 --reference-doc으로 사용하여 양식을 보존합니다.

pypandoc-hwpx 0.1.1 버그 후처리 패치 포함:
  - 빈 hp:subList에 빈 paragraph 주입 (한글 크래시 방지)
  - hp:tbl의 borderFillIDRef를 셀과 동일하게 통일
"""
import os
import re
import sys
import argparse
import subprocess
import zipfile
import io


def _patch_hwpx(output_path):
    """pypandoc-hwpx 출력물의 알려진 버그를 후처리로 수정.

    1) 빈 hp:subList → 빈 paragraph를 주입 (한글이 빈 subList에서 크래시)
    2) hp:tbl의 borderFillIDRef를 첫 번째 셀의 값과 동일하게 교체
       (pypandoc-hwpx가 "3"으로 하드코딩하지만 reference-doc 사용 시
        ID 3이 테이블 보더가 아닌 다른 용도의 borderFill일 수 있음)
    """
    with open(output_path, 'rb') as f:
        original_bytes = f.read()

    z_in = zipfile.ZipFile(io.BytesIO(original_bytes), 'r')

    # Contents/section*.xml 파일 찾기
    section_pattern = re.compile(r'Contents/section\d+\.xml')
    section_files = [name for name in z_in.namelist() if section_pattern.match(name)]

    if not section_files:
        z_in.close()
        return False

    patched_sections = {}
    any_patched = False

    # 각 섹션 파일에 대해 패치 적용
    for section_file in section_files:
        sec_xml = z_in.read(section_file).decode('utf-8')
        section_patched = False

        # 패치 1: 빈 hp:subList에 빈 paragraph 주입
        empty_sublist_pattern = r'(<hp:subList[^>]*>)(</hp:subList>)'
        if re.search(empty_sublist_pattern, sec_xml):
            # 문서에서 사용하는 paraPrIDRef 찾기 (첫 번째 hp:p에서 추출)
            para_pr_match = re.search(r'<hp:p\s+paraPrIDRef="(\d+)"', sec_xml)
            para_pr_id = para_pr_match.group(1) if para_pr_match else "0"

            empty_para = (
                f'<hp:p paraPrIDRef="{para_pr_id}" styleIDRef="0"'
                f' pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run></hp:p>'
            )
            sec_xml = re.sub(empty_sublist_pattern, r'\1' + empty_para + r'\2', sec_xml)
            section_patched = True

        # 패치 2: hp:tbl의 borderFillIDRef를 셀과 통일
        # 첫 번째 hp:tc의 borderFillIDRef를 찾아서 hp:tbl에도 적용
        tbl_pattern = r'(<hp:tbl[^>]*?)borderFillIDRef="(\d+)"([^>]*>)'
        tc_bf_pattern = r'<hp:tc[^>]*borderFillIDRef="(\d+)"'

        for tbl_match in re.finditer(tbl_pattern, sec_xml):
            tbl_bf_id = tbl_match.group(2)
            # 이 테이블 이후 첫 번째 tc의 borderFillIDRef 찾기
            rest = sec_xml[tbl_match.end():]
            tc_match = re.search(tc_bf_pattern, rest)
            if tc_match and tc_match.group(1) != tbl_bf_id:
                cell_bf_id = tc_match.group(1)
                old = tbl_match.group(0)
                new = old.replace(f'borderFillIDRef="{tbl_bf_id}"',
                                  f'borderFillIDRef="{cell_bf_id}"')
                sec_xml = sec_xml.replace(old, new, 1)
                section_patched = True

        if section_patched:
            patched_sections[section_file] = sec_xml
            any_patched = True

    if not any_patched:
        z_in.close()
        return False

    # 수정된 섹션 파일들을 ZIP에 다시 쓰기
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            if item.filename in patched_sections:
                z_out.writestr(item.filename, patched_sections[item.filename].encode('utf-8'))
            elif item.filename == 'mimetype':
                z_out.writestr(item, z_in.read(item.filename),
                               compress_type=zipfile.ZIP_STORED)
            else:
                z_out.writestr(item, z_in.read(item.filename))
    z_in.close()

    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())

    return True


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

    # pypandoc-hwpx 버그 후처리 패치
    if _patch_hwpx(output_path):
        print("후처리 패치 적용: 빈 셀/borderFill 수정")

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
