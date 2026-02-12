"""
app.py - HWPX 편집 대시보드 (Streamlit MVP)

HWPX 파일을 업로드하여 마크다운으로 변환하고 편집한 뒤,
다시 HWPX로 내보내는 웹 기반 편집기
"""
import streamlit as st
import os
import tempfile
from pathlib import Path

# 페이지 설정
st.set_page_config(
    page_title="HWPX 편집 대시보드",
    page_icon="📝",
    layout="wide"
)

# 서비스 import
from services.pipeline_service import PipelineService


def format_file_size(size_bytes):
    """파일 크기를 읽기 쉬운 형식으로 변환"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def init_session_state():
    """세션 상태 초기화"""
    if 'service' not in st.session_state:
        st.session_state.service = PipelineService()
    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    if 'original_md' not in st.session_state:
        st.session_state.original_md = None
    if 'edited_md' not in st.session_state:
        st.session_state.edited_md = None
    if 'hwpx_info' not in st.session_state:
        st.session_state.hwpx_info = None
    if 'temp_hwpx_path' not in st.session_state:
        st.session_state.temp_hwpx_path = None
    if 'conversion_done' not in st.session_state:
        st.session_state.conversion_done = False


def main():
    init_session_state()

    st.title("📝 HWPX 편집 대시보드")
    st.markdown("HWPX 파일을 마크다운으로 변환하여 편집하고, 다시 HWPX로 저장합니다.")

    # 사이드바 - 파일 업로드
    with st.sidebar:
        st.header("파일 업로드")

        uploaded_file = st.file_uploader(
            "HWPX 파일 선택",
            type=['hwpx'],
            help="편집할 HWPX 파일을 업로드하세요"
        )

        # linesegarray 제거 옵션
        strip_lineseg = st.checkbox(
            "텍스트 겹침 방지 (linesegarray 제거)",
            value=True,
            help="텍스트가 겹쳐 보이는 문제를 방지합니다"
        )

        if uploaded_file is not None:
            # 파일이 변경되었는지 확인
            if st.session_state.uploaded_file != uploaded_file.name:
                st.session_state.uploaded_file = uploaded_file.name
                st.session_state.conversion_done = False
                st.session_state.original_md = None
                st.session_state.edited_md = None

                # 임시 파일로 저장
                with tempfile.NamedTemporaryFile(delete=False, suffix='.hwpx') as tmp:
                    tmp.write(uploaded_file.getvalue())
                    st.session_state.temp_hwpx_path = tmp.name

                # linesegarray 제거
                if strip_lineseg:
                    with st.spinner("텍스트 겹침 방지 처리 중..."):
                        result = st.session_state.service.strip_lineseg(st.session_state.temp_hwpx_path)
                        if not result['success']:
                            st.warning(f"linesegarray 제거 실패: {result['message']}")

                # 파일 정보 추출
                st.session_state.hwpx_info = st.session_state.service.get_hwpx_info(
                    st.session_state.temp_hwpx_path
                )

            # 변환 버튼
            if st.button("🔄 마크다운으로 변환", type="primary", use_container_width=True):
                with st.spinner("변환 중..."):
                    try:
                        result = st.session_state.service.convert_to_markdown(
                            st.session_state.temp_hwpx_path
                        )
                        st.session_state.original_md = result['md_content']
                        st.session_state.edited_md = result['md_content']
                        st.session_state.conversion_done = True
                        st.success("✅ 변환 완료!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 변환 실패: {str(e)}")

            # 파일 정보 표시
            if st.session_state.hwpx_info:
                st.divider()
                st.subheader("파일 정보")
                info = st.session_state.hwpx_info

                if 'error' in info:
                    st.error(f"정보 추출 실패: {info['error']}")
                else:
                    st.metric("파일 크기", format_file_size(info['file_size']))
                    st.metric("섹션 수", f"{info['section_count']}개")
                    st.metric("테이블 수", f"{info['table_count']}개")
                    st.metric("문단 수", f"{info['paragraph_count']}개")

    # 메인 영역
    if not st.session_state.conversion_done:
        # 변환 전 안내 화면
        st.info("👈 왼쪽 사이드바에서 HWPX 파일을 업로드하고 '마크다운으로 변환' 버튼을 클릭하세요.")

        st.markdown("### 사용 방법")
        st.markdown("""
        1. **파일 업로드**: 사이드바에서 HWPX 파일을 선택합니다
        2. **마크다운 변환**: '마크다운으로 변환' 버튼을 클릭합니다
        3. **텍스트 편집**: '편집' 탭에서 마크다운 텍스트를 수정합니다
        4. **HWPX 생성**: '변경사항 & 다운로드' 탭에서 HWPX 파일을 생성하고 다운로드합니다
        """)

        st.markdown("### 주의사항")
        st.markdown("""
        - 테이블 구조(행/열 수) 변경은 지원하지 않습니다
        - 텍스트 내용만 편집 가능합니다
        - 표 서식은 원본 그대로 보존됩니다
        """)

    else:
        # 변환 후 편집 화면
        tab1, tab2 = st.tabs(["📝 편집", "💾 변경사항 & 다운로드"])

        with tab1:
            st.header("마크다운 편집")

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("편집기")
                edited_text = st.text_area(
                    "마크다운 텍스트",
                    value=st.session_state.edited_md,
                    height=600,
                    help="마크다운 문법으로 편집하세요",
                    label_visibility="collapsed"
                )

                # 편집 내용 저장
                if edited_text != st.session_state.edited_md:
                    st.session_state.edited_md = edited_text

            with col2:
                st.subheader("미리보기")
                # 스크롤 가능한 컨테이너
                with st.container(height=600):
                    st.markdown(st.session_state.edited_md)

        with tab2:
            st.header("변경사항 및 다운로드")

            # 변경사항 분석
            if st.session_state.original_md and st.session_state.edited_md:
                changes = st.session_state.service.analyze_changes(
                    st.session_state.original_md,
                    st.session_state.edited_md
                )

                # 변경사항 표시
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("전체 테이블", f"{changes['total_tables']}개")
                with col2:
                    st.metric("변경된 셀", f"{changes['table_changes']}개")
                with col3:
                    st.metric("전체 문단", f"{changes['total_paragraphs']}개")
                with col4:
                    st.metric("변경된 문단", f"{changes['paragraph_changes']}개")

                st.divider()

                # HWPX 생성 버튼
                if st.button("🔨 HWPX 생성", type="primary", use_container_width=True):
                    with st.spinner("HWPX 파일 생성 중..."):
                        try:
                            # 임시 마크다운 파일 저장
                            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                                            suffix='.md', delete=False) as tmp_md:
                                tmp_md.write(st.session_state.edited_md)
                                tmp_md_path = tmp_md.name

                            # 출력 HWPX 경로
                            output_hwpx = tempfile.NamedTemporaryFile(delete=False, suffix='.hwpx').name

                            # smart_replace 실행
                            result = st.session_state.service.smart_replace(
                                st.session_state.temp_hwpx_path,
                                tmp_md_path,
                                output_hwpx
                            )

                            if result['success']:
                                st.success("✅ HWPX 생성 완료!")

                                # 파일 읽기
                                with open(result['output_path'], 'rb') as f:
                                    hwpx_bytes = f.read()

                                # 다운로드 버튼
                                original_name = Path(st.session_state.uploaded_file).stem
                                st.download_button(
                                    label="📥 HWPX 다운로드",
                                    data=hwpx_bytes,
                                    file_name=f"{original_name}_edited.hwpx",
                                    mime="application/octet-stream",
                                    use_container_width=True
                                )

                                # 임시 파일 정리
                                try:
                                    os.unlink(tmp_md_path)
                                except:
                                    pass
                            else:
                                st.error(f"❌ {result['message']}")

                        except Exception as e:
                            st.error(f"❌ 생성 실패: {str(e)}")
                            import traceback
                            with st.expander("오류 상세 정보"):
                                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
