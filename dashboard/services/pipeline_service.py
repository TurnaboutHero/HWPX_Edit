"""
pipeline_service.py - Pipeline 모듈 래퍼 서비스

Streamlit 앱에서 pipeline 모듈을 사용하기 위한 서비스 레이어
"""
import os
import sys
import tempfile
import zipfile
from pathlib import Path

# pipeline 모듈을 import하기 위해 상위 디렉토리를 sys.path에 추가
DASHBOARD_DIR = Path(__file__).parent.parent
PROJECT_ROOT = DASHBOARD_DIR.parent
PIPELINE_DIR = PROJECT_ROOT / "pipeline"

if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from hwpx_to_md import convert_hwpx_to_md, HwpxToMarkdown
from smart_replace import (
    smart_replace,
    strip_linesegarray,
    parse_markdown_tables,
    parse_markdown_paragraphs,
    collect_text_growth_warnings,
    collect_protected_markup_changes,
    format_growth_warning,
    format_protected_markup_change,
)


class PipelineService:
    """Pipeline 기능을 Streamlit 앱에서 사용하기 위한 서비스 클래스"""

    def __init__(self):
        self.temp_dir = None

    def convert_to_markdown(self, hwpx_path, output_dir=None):
        """HWPX 파일을 마크다운으로 변환

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_dir: 출력 디렉토리 (None이면 임시 디렉토리 사용)

        Returns:
            dict: {
                'md_path': 마크다운 파일 경로,
                'md_content': 마크다운 텍스트,
                'image_count': 추출된 이미지 수,
                'images_dir': 이미지 디렉토리 경로
            }
        """
        if output_dir is None:
            if self.temp_dir is None:
                self.temp_dir = tempfile.mkdtemp(prefix="hwpx_edit_")
            output_dir = self.temp_dir

        # 출력 경로 생성
        base_name = Path(hwpx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}.md")

        # 변환 실행
        converter = HwpxToMarkdown(hwpx_path, output_dir=output_dir, extract_images=True)
        md_content = converter.convert()

        # 파일로 저장
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

        return {
            'md_path': output_path,
            'md_content': md_content,
            'image_count': len(converter.image_map),
            'images_dir': converter.images_dir if converter.image_map else None
        }

    def smart_replace(self, original_hwpx, edited_md_path, output_hwpx,
                      strip_lineseg=True, allow_layout_risk=False):
        """편집된 마크다운을 원본 HWPX에 반영

        Args:
            original_hwpx: 원본 HWPX 파일 경로
            edited_md_path: 편집된 마크다운 파일 경로
            output_hwpx: 출력 HWPX 파일 경로

        Returns:
            dict: {
                'success': bool,
                'output_path': 출력 파일 경로,
                'message': 결과 메시지
            }
        """
        try:
            result_path = smart_replace(
                original_hwpx,
                edited_md_path,
                output_hwpx,
                strip_lineseg=strip_lineseg,
                allow_layout_risk=allow_layout_risk,
            )
            return {
                'success': True,
                'output_path': result_path,
                'message': '변환 완료'
            }
        except Exception as e:
            return {
                'success': False,
                'output_path': None,
                'message': f'변환 실패: {str(e)}'
            }

    def analyze_changes(self, original_md, edited_md):
        """원본과 편집본 마크다운의 변경사항 분석

        Args:
            original_md: 원본 마크다운 텍스트
            edited_md: 편집된 마크다운 텍스트

        Returns:
            dict: {
                'table_changes': 변경된 테이블 셀 수,
                'paragraph_changes': 변경된 문단 수,
                'total_tables': 전체 테이블 수,
                'total_paragraphs': 전체 문단 수
            }
        """
        # 테이블 분석
        orig_tables = parse_markdown_tables(original_md)
        edited_tables = parse_markdown_tables(edited_md)

        table_changes = 0
        orig_cells = []
        edited_cells = []
        for i in range(min(len(orig_tables), len(edited_tables))):
            orig = orig_tables[i]
            edited = edited_tables[i]

            # 셀 비교
            for row_idx in range(min(len(orig['cells']), len(edited['cells']))):
                for col_idx in range(min(len(orig['cells'][row_idx]), len(edited['cells'][row_idx]))):
                    orig_cell = orig['cells'][row_idx][col_idx]
                    edited_cell = edited['cells'][row_idx][col_idx]
                    orig_cells.append(orig_cell)
                    edited_cells.append(edited_cell)
                    if orig_cell.strip() != edited_cell.strip():
                        table_changes += 1

        # 문단 분석
        orig_paras = parse_markdown_paragraphs(original_md)
        edited_paras = parse_markdown_paragraphs(edited_md)

        para_changes = 0
        for i in range(min(len(orig_paras), len(edited_paras))):
            if orig_paras[i].strip() != edited_paras[i].strip():
                para_changes += 1

        layout_warnings = []
        layout_warnings.extend(
            format_growth_warning(w)
            for w in collect_text_growth_warnings(
                orig_cells,
                edited_cells,
                kind='셀',
            )
        )
        layout_warnings.extend(
            format_growth_warning(w)
            for w in collect_text_growth_warnings(
                orig_paras,
                edited_paras,
                kind='문단',
            )
        )
        protected_changes = [
            format_protected_markup_change(change)
            for change in collect_protected_markup_changes(original_md, edited_md)
        ]

        return {
            'table_changes': table_changes,
            'paragraph_changes': para_changes,
            'total_tables': len(orig_tables),
            'total_paragraphs': len(orig_paras),
            'edited_tables': len(edited_tables),
            'edited_paragraphs': len(edited_paras),
            'table_count_match': len(orig_tables) == len(edited_tables),
            'paragraph_count_match': len(orig_paras) == len(edited_paras),
            'layout_warnings': layout_warnings,
            'protected_changes': protected_changes,
        }

    def strip_lineseg(self, hwpx_path, output_path=None):
        """HWPX에서 linesegarray 제거 (텍스트 겹침 방지)

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_path: 출력 HWPX 파일 경로 (None이면 원본 덮어쓰기)

        Returns:
            dict: {'success': bool, 'message': str}
        """
        try:
            if output_path is None:
                output_path = hwpx_path

            if output_path != hwpx_path:
                import shutil
                shutil.copy2(hwpx_path, output_path)

            removed = strip_linesegarray(output_path)

            return {
                'success': True,
                'message': f'linesegarray 제거 완료: {removed}개'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'제거 실패: {str(e)}'
            }

    def get_hwpx_info(self, hwpx_path):
        """HWPX 파일 정보 추출

        Args:
            hwpx_path: HWPX 파일 경로

        Returns:
            dict: {
                'file_size': 파일 크기 (bytes),
                'table_count': 테이블 수,
                'paragraph_count': 문단 수,
                'section_count': 섹션 수
            }
        """
        try:
            file_size = os.path.getsize(hwpx_path)

            # 간단한 정보 추출 (전체 변환 없이)
            with zipfile.ZipFile(hwpx_path, 'r') as z:
                section_files = [name for name in z.namelist() if name.startswith('Contents/section') and name.endswith('.xml')]
                section_count = len(section_files)

                # 첫 번째 섹션에서 테이블/문단 수 추출
                table_count = 0
                para_count = 0

                if section_files:
                    for section_file in section_files:
                        data = z.read(section_file).decode('utf-8')
                        table_count += data.count('<hp:tbl ')
                        para_count += data.count('<hp:p ')

            return {
                'file_size': file_size,
                'table_count': table_count,
                'paragraph_count': para_count,
                'section_count': section_count
            }
        except Exception as e:
            return {
                'file_size': 0,
                'table_count': 0,
                'paragraph_count': 0,
                'section_count': 0,
                'error': str(e)
            }

    def cleanup(self):
        """임시 파일 정리"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            import shutil
            try:
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
            except Exception:
                pass
