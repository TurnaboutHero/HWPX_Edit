# HWPX_Edit TODO

## 다음 세션에서 할 일

### 1. pypandoc-hwpx + smart_replace 실제 결합 (Strategy A: Diff-Aware Router)
- 현재 `convert.py auto`는 변경 감지만 하고 항상 smart_replace만 호출
- **구조 변경**(테이블 추가/삭제, 행/열 변경)시 pypandoc-hwpx로 새 HWPX 생성 후 원본과 병합하는 로직 필요
- `pipeline/md_to_hwpx.py`의 `convert_md_to_hwpx()` 활용
- 라우팅 로직: 텍스트만 변경 → smart_replace / 구조 변경 → pypandoc-hwpx + 서식 병합

### 2. Streamlit 대시보드 테스트
- `streamlit run dashboard/app.py`로 실행 후 실제 HWPX 업로드 → 편집 → 다운로드 워크플로 확인
- PipelineService 각 메서드 동작 확인 (get_hwpx_info, convert_to_markdown, smart_replace, strip_lineseg, analyze_changes)

### 3. 문단 파싱 휴리스틱 개선
- 원본 XML 156개 문단 vs 마크다운 73개 문단 불일치 문제
- `parse_markdown_paragraphs()`와 `extract_xml_paragraphs()` 매칭률 향상 필요

### 4. Codex 리뷰 Medium 이슈
- blockquote/1x1 테이블 혼동
- escaped pipe (`\|`) 파싱
- 테스트 커버리지 확대

## 완료된 항목
- [x] smart_replace Critical/High 이슈 3건 수정 (bbfce68)
- [x] convert.py auto 서브커맨드 + Streamlit 대시보드 MVP (a6ccf57)
- [x] cp949 이모지 인코딩 문제 수정
- [x] linesegarray 제거 (텍스트 겹침 방지)
- [x] 신청서 내용 강화 (시장 데이터, 매출 근거, 경쟁 차별성, 투자 타임라인)
