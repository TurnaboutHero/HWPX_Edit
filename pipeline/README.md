# HWPX ↔ Markdown 변환 파이프라인

한컴오피스 한글 HWPX 문서를 마크다운으로 변환하고, 편집 후 다시 HWPX로 복원하는 도구입니다.

## 주요 특징

- **HWPX → Markdown**: 표, 이미지, 서식, 제목, 양식, 수식 등 완벽 지원
- **스마트 교체**: 원본 구조 100% 보존하며 텍스트만 반영 (권장)
- **Markdown → HWPX**: pypandoc-hwpx 기반 완전 변환
- **다중 섹션 지원**: section0~N.xml 자동 처리
- **OWPML 2024 호환**: 2011/2024 네임스페이스 자동 감지

## 설치

### 필수 의존성

| 패키지 | 버전 | 용도 |
|--------|------|------|
| Python | 3.13+ | 런타임 |
| Pandoc | 3.9 | pypandoc-hwpx 내부 사용 |
| pypandoc-hwpx | 0.1.1 | md → hwpx 변환 |
| pypandoc | 1.16.2 | Pandoc Python 래퍼 |
| lxml | 6.0.2+ | HWPX XML 파싱 |
| Pillow | 12.0.0+ | BMP → PNG 이미지 변환 |

### 설치 방법

```bash
# 1. Pandoc 설치 (https://pandoc.org/installing.html)
# Windows: choco install pandoc
# macOS: brew install pandoc
# Ubuntu: sudo apt install pandoc

# 2. Python 패키지 설치
pip install lxml Pillow pypandoc pypandoc-hwpx
```

## 빠른 시작

### HWPX → Markdown

```bash
python convert.py to-md 문서.hwpx
# 출력: 문서.md, images/ 폴더
```

### Markdown → HWPX (스마트 교체 — 권장)

```bash
python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx
```

### Markdown → HWPX (pypandoc 경유)

```bash
python convert.py to-hwpx 문서.md -r 원본.hwpx -o 최종본.hwpx
```

## 사용법 상세

### 1. HWPX → Markdown (`hwpx_to_md.py`)

HWPX 파일의 모든 요소를 마크다운으로 변환합니다.

```bash
# 기본 사용
python convert.py to-md 신청서.hwpx

# 출력 경로 지정
python convert.py to-md 신청서.hwpx -o output/문서.md

# 이미지 추출 비활성화
python convert.py to-md 신청서.hwpx --no-images
```

**지원 기능**:
- 표 (`hp:tbl`) → 마크다운 테이블
- 이미지 (`hp:pic`) → `![ref](images/...)` (BMP→PNG 자동 변환)
- 서식 (`hp:run`) → `**굵게**`, `*기울임*`, `~~취소선~~`
- 제목 (`hh:heading`) → `#` ~ `######`
- 수식 (`hp:equation`) → `$$ LaTeX $$`
- 머리글/꼬리글 (`hp:header/footer`) → `<!-- 머리글: ... -->`
- 양식 (`hp:checkBtn/radioBtn/comboBox`) → `[ ]`, `(o)`, `[콤보: ...]`
- 덧말 (`hp:dutmal`) → `<ruby>본말<rt>닷말</rt></ruby>`
- 글맵시 (`hp:textart`) → `> [글맵시] text`
- OLE 개체 (`hp:ole`) → `<!-- OLE: ... -->`
- 변경 추적 (`hp:deleteBegin/End`) → `~~삭제~~`
- 하이퍼링크 (`hp:fieldBegin HYPERLINK`) → `[text](url)`
- 다단 레이아웃 (`hp:colPr`) → `<!-- [N단 레이아웃] -->`
- 다중 섹션 자동 처리 (section0~N.xml → `---` 구분자)

**출력 구조**:
```
문서.md                    # 변환된 마크다운
images/                    # 추출된 이미지
  ├── image1.png
  └── image2.png
template_info.json         # 페이지 설정 메타데이터
```

### 2. 스마트 교체 (`smart_replace.py`) — 권장

원본 HWPX 구조를 100% 보존하면서 편집된 마크다운의 텍스트만 반영합니다.

```bash
# 기본 사용
python convert.py smart 원본.hwpx 편집된.md

# 출력 경로 지정
python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx
```

**주요 장점**:
- ✅ 표 서식 완벽 보존 (셀 병합, 너비, 높이, 배경색, 테두리)
- ✅ 원본 XML 바이트 수준 보존 (lxml 직렬화 우회)
- ✅ pypandoc-hwpx 버그 우회 (빈 셀 크래시, borderFill 불일치)
- ✅ 다중 섹션 지원 (section0~N.xml 전체 처리)
- ✅ OWPML 2024 네임스페이스 자동 감지

**제한사항**:
- ❌ 테이블 행/열 추가/삭제 불가 (구조 변경 불가)
- ❌ 문단 추가/삭제 불가
- ✅ 기존 셀/문단의 텍스트 내용 변경만 가능

**권장 워크플로**:
```bash
# 1. HWPX → 마크다운
python convert.py to-md 원본.hwpx -o work/문서.md

# 2. AI나 텍스트 에디터로 work/문서.md 편집 (텍스트만 변경)

# 3. 스마트 교체로 원본에 반영
python convert.py smart 원본.hwpx work/문서.md -o 최종본.hwpx
```

### 3. Markdown → HWPX (`md_to_hwpx.py`)

pypandoc-hwpx를 사용하여 마크다운을 HWPX로 완전 변환합니다.

```bash
# 기본 사용 (기본 템플릿 사용)
python convert.py to-hwpx 사업계획서.md

# 원본 양식 보존
python convert.py to-hwpx 사업계획서.md -r 원본양식.hwpx -o 최종본.hwpx
```

**`--reference-doc` 옵션**:
- 원본 HWPX의 스타일 정의 (`header.xml`) 복사
- 페이지 크기/여백/방향 유지
- 폰트/문단속성 유지
- 기존 이미지 파일 유지

**자동 버그 패치**:
- 빈 마크다운 셀 → 빈 `hp:p` 주입 (한글 크래시 방지)
- `hp:tbl`의 `borderFillIDRef`를 셀과 통일 (테두리 스타일 보존)

## 워크플로 예시

### AI 편집 워크플로 (스마트 교체)

```bash
# 1. HWPX를 마크다운으로 변환
python convert.py to-md 신청서.hwpx -o work/신청서.md

# 2. AI로 work/신청서.md 편집 (텍스트만 변경)
# 예: ChatGPT, Claude에 업로드하여 맞춤법 교정, 내용 개선

# 3. 원본 구조 유지하며 텍스트만 반영
python convert.py smart 신청서.hwpx work/신청서.md -o 제출용_신청서.hwpx
```

### 새 문서 작성 워크플로 (pypandoc)

```bash
# 1. 마크다운으로 문서 작성
vim 사업계획서.md

# 2. 원본 양식 기반 HWPX 생성
python convert.py to-hwpx 사업계획서.md -r 양식템플릿.hwpx -o 최종본.hwpx
```

### 대량 변환 워크플로

```bash
# 여러 HWPX 파일을 일괄 마크다운으로 변환
for file in *.hwpx; do
    python convert.py to-md "$file" -o "output/${file%.hwpx}.md"
done
```

## 지원 기능

| 카테고리 | HWPX → MD | 스마트 교체 | MD → HWPX |
|---------|-----------|-------------|-----------|
| **텍스트 서식** | | | |
| 굵기/기울임/취소선 | ✅ | ✅ | ✅ |
| 제목 레벨 (H1~H6) | ✅ | ✅ | ✅ |
| 글꼴/색상 | ⚠️ 메타데이터만 | ✅ | ⚠️ 제한적 |
| **구조 요소** | | | |
| 표 (기본) | ✅ | ✅ | ✅ |
| 표 (셀 병합) | ✅ | ✅ | ❌ |
| 표 (셀 너비/높이) | ⚠️ 메타데이터만 | ✅ | ❌ |
| 표 (배경색/테두리) | ⚠️ 메타데이터만 | ✅ | ❌ |
| 이미지 | ✅ | ✅ | ✅ |
| 다중 섹션 | ✅ | ✅ | ⚠️ 단일 섹션 |
| **고급 기능** | | | |
| 수식 | ✅ | ✅ | ⚠️ 제한적 |
| 머리글/꼬리글 | ✅ | ✅ | ❌ |
| 각주/미주 | ✅ | ✅ | ✅ |
| 하이퍼링크 | ✅ | ✅ | ✅ |
| **양식 개체** | | | |
| 체크박스/라디오 | ✅ | ✅ | ❌ |
| 콤보박스/입력란 | ✅ | ✅ | ❌ |
| 버튼 | ✅ | ✅ | ❌ |
| **특수 요소** | | | |
| 덧말(ruby) | ✅ | ✅ | ❌ |
| 글맵시(TextArt) | ✅ | ✅ | ❌ |
| OLE 개체 | ✅ | ✅ | ❌ |
| 변경 추적 | ✅ | ✅ | ❌ |
| 다단 레이아웃 | ✅ | ✅ | ❌ |

**범례**:
- ✅ 완전 지원
- ⚠️ 부분 지원 (메타데이터만 보존, 시각적 효과 미반영)
- ❌ 미지원

## 제한사항

### 스마트 교체 vs pypandoc 비교

| 항목 | 스마트 교체 | pypandoc (to-hwpx) |
|------|-------------|-------------------|
| **장점** | | |
| 원본 구조 보존 | ✅ 완벽 | ❌ 손실 |
| 셀 병합 보존 | ✅ | ❌ 해제됨 |
| 셀 너비/높이 보존 | ✅ | ❌ 균등화 |
| 셀 배경색/테두리 | ✅ | ❌ 소실 |
| XML 바이트 보존 | ✅ | ❌ 재생성 |
| **단점** | | |
| 텍스트 내용 변경 | ✅ | ✅ |
| 새 표/행/열 추가 | ❌ | ✅ |
| 문단 추가/삭제 | ❌ | ✅ |
| 구조 변경 | ❌ | ✅ |

**권장 사항**:
- **기존 문서 편집** (내용만 수정) → **스마트 교체** 사용
- **새 문서 작성** (구조 변경 필요) → **to-hwpx** 사용

### 알려진 문제

1. **pypandoc-hwpx 0.1.1 버그**:
   - 빈 마크다운 셀 → 한글 크래시 (자동 패치됨)
   - `borderFillIDRef` 하드코딩 (자동 패치됨)

2. **마크다운 한계**:
   - 셀 병합 (colspan/rowspan) 미지원
   - 셀 배경색/테두리 스타일 미지원
   - 표 안의 표 (중첩 테이블) 미지원

3. **HWPX 복잡도**:
   - 매우 복잡한 레이아웃은 마크다운 변환 시 단순화됨
   - 일부 서식 (장평/자간/그림자 등)은 메타데이터로만 보존

## 기술 문서

더 자세한 기술 정보는 다음 문서를 참고하세요:

- **NOTES.md**: HWPX 구조, XML 네임스페이스, 파이프라인 작동 원리 상세
- **convert.py**: 통합 CLI 사용법 및 옵션
- **hwpx_to_md.py**: HWPX → Markdown 변환 로직
- **smart_replace.py**: 스마트 교체 알고리즘
- **md_to_hwpx.py**: Markdown → HWPX 변환 및 버그 패치

## 라이선스

MIT License

이 프로젝트는 다음 오픈소스 프로젝트를 사용합니다:
- [pypandoc-hwpx](https://github.com/msjang/pypandoc-hwpx) (MIT)
- [Pandoc](https://pandoc.org/) (GPL v2)
- [lxml](https://lxml.de/) (BSD)

## 기여

버그 리포트, 기능 제안, 풀 리퀘스트를 환영합니다.

### 테스트

```bash
# hwpxlib 테스트 파일로 검증 (47개 파일)
python test_pipeline.py

# 개별 테스트
python convert.py to-md test_data/SimpleTable.hwpx
python convert.py smart test_data/SimpleTable.hwpx test_data/SimpleTable.md
```

## FAQ

### Q1. 왕복 변환(roundtrip) 시 서식이 깨집니다.

**A**: `smart` 명령어를 사용하세요. pypandoc 대신 원본 XML을 직접 수정하여 서식을 100% 보존합니다.

```bash
python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx
```

### Q2. 변환된 HWPX가 한글에서 열리지 않습니다.

**A**: pypandoc-hwpx 0.1.1의 빈 셀 버그 때문입니다. 이 파이프라인은 자동으로 패치를 적용하므로, `md_to_hwpx.py`를 통해 변환하면 문제가 해결됩니다.

### Q3. 표의 셀 병합이 사라집니다.

**A**: 마크다운은 셀 병합(colspan/rowspan)을 지원하지 않습니다. 기존 문서 편집 시에는 `smart` 명령어를 사용하면 원본 병합이 보존됩니다.

### Q4. 이미지가 변환되지 않습니다.

**A**: `--no-images` 옵션을 사용하지 않았는지 확인하세요. BMP 이미지는 자동으로 PNG로 변환됩니다.

### Q5. OWPML 2024 문서도 지원하나요?

**A**: 네, 자동으로 감지하여 처리합니다. 2011/2024 네임스페이스 모두 지원합니다.

## 지원

문의사항이나 버그 리포트는 GitHub Issues에 등록해주세요.
