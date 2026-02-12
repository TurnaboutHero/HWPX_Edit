# 파이프라인 기술 문서

## 1. HWPX 파일 구조 분석

HWPX는 한컴오피스 한글의 개방형 문서 포맷으로, **ZIP 압축된 XML 파일 묶음**이다.

```
파일.hwpx (ZIP)
├── mimetype                    # MIME 타입 선언
├── version.xml                 # 버전 정보
├── Contents/
│   ├── header.xml              # 스타일 정의 (폰트, 문단속성, 글자속성, 제목레벨)
│   ├── section0.xml            # 본문 내용 (문단, 표, 이미지 참조)
│   └── content.hpf             # 패키지 매니페스트 (파일 목록, 제목)
├── BinData/
│   ├── image1.bmp              # 이미지 바이너리 (BMP, JPG 등)
│   ├── image2.jpg
│   └── ...
├── Preview/
│   ├── PrvText.txt             # 텍스트 미리보기
│   └── PrvImage.png            # 이미지 미리보기
├── settings.xml                # 문서 설정
└── META-INF/                   # 메타데이터
    ├── container.xml
    ├── container.rdf
    └── manifest.xml
```

## 2. 핵심 XML 네임스페이스

| 접두사 | URI | 용도 |
|--------|-----|------|
| `hp` | `http://www.hancom.co.kr/hwpml/2011/paragraph` | 문단, 텍스트런, 표, 이미지 |
| `hh` | `http://www.hancom.co.kr/hwpml/2011/head` | 헤더(스타일 정의) |
| `hc` | `http://www.hancom.co.kr/hwpml/2011/core` | 이미지 데이터, 채우기 등 |
| `hs` | `http://www.hancom.co.kr/hwpml/2011/section` | 섹션 루트 |

**OWPML 2024 네임스페이스** (자동 감지):

| 접두사 | URI | 용도 |
|--------|-----|------|
| `hp` | `http://www.owpml.org/owpml/2024/paragraph` | 문단, 텍스트런, 표, 이미지 |
| `hh` | `http://www.owpml.org/owpml/2024/head` | 헤더(스타일 정의) |
| `hc` | `http://www.owpml.org/owpml/2024/core` | 이미지 데이터, 채우기 등 |
| `hs` | `http://www.owpml.org/owpml/2024/body` | 섹션(body로 변경) |
| `hv` | `http://www.owpml.org/owpml/2024/version` | 버전 정보 |
| `hm` | `http://www.owpml.org/owpml/2024/master-page` | 마스터 페이지 |
| `hhs` | `http://www.owpml.org/owpml/2024/history` | 변경 이력 |

## 3. section0.xml 본문 구조

```xml
<hs:sec>
  <hp:p paraPrIDRef="73" styleIDRef="0">     <!-- 문단 -->
    <hp:run charPrIDRef="9">                  <!-- 텍스트 런 (서식 단위) -->
      <hp:t>본문 텍스트</hp:t>                <!-- 실제 텍스트 -->
    </hp:run>
    <hp:run charPrIDRef="13">
      <hp:t>굵은 텍스트</hp:t>                <!-- charPrIDRef로 서식 참조 -->
    </hp:run>
  </hp:p>
</hs:sec>
```

### 주요 요소

- **`hp:p`** - 문단. `paraPrIDRef`로 문단속성, `styleIDRef`로 스타일 참조
- **`hp:run`** - 텍스트 런. `charPrIDRef`로 글자속성(굵기, 기울임 등) 참조
- **`hp:t`** - 실제 텍스트 노드. 내부에 `hp:lineBreak` 등 포함 가능
- **`hp:tbl`** - 표. `rowCnt`, `colCnt` 속성으로 크기 정의
  - `hp:tr` → `hp:tc` (셀) → `hp:cellAddr` (위치) + `hp:cellSpan` (병합)
  - 셀 내부에 `hp:subList` → `hp:p` 형태로 내용 포함
- **`hp:pic`** - 이미지. 내부 `hc:img binaryItemIDRef="image1"`로 BinData 참조

### 표(Table) 구조 예시

```xml
<hp:tbl rowCnt="3" colCnt="4" cellSpacing="0">
  <hp:tr>
    <hp:tc borderFillIDRef="22">
      <hp:subList>
        <hp:p><hp:run><hp:t>셀 내용</hp:t></hp:run></hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="2" rowSpan="1"/>    <!-- 2열 병합 -->
      <hp:cellSz width="23000" height="1000"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

### 이미지 참조 구조

```xml
<hp:pic id="2082278702">
  <hc:img binaryItemIDRef="image1"/>    <!-- BinData/image1.bmp 참조 -->
  <hp:imgDim dimwidth="127260" dimheight="77580"/>
</hp:pic>
```

## 4. header.xml 스타일 정의

### 글자속성 (charPr)

```xml
<hh:charPr id="9">
  <hh:bold/>                                    <!-- 존재하면 굵게 -->
  <hh:italic/>                                  <!-- 존재하면 기울임 -->
  <hh:underline type="BOTTOM" shape="SOLID"/>   <!-- type="NONE"이면 없음 -->
  <hh:strikeout shape="NONE"/>                  <!-- shape="NONE"이면 없음! -->
</hh:charPr>
```

**주의**: `strikeout`은 `type`이 아니라 **`shape`** 속성으로 판별.
`shape="NONE"`이면 취소선 없음, 다른 값이면 취소선 있음.

### 문단속성 (paraPr) & 제목 레벨

```xml
<hh:paraPr id="23">
  <hh:heading type="OUTLINE" level="0"/>   <!-- level 0 = 제목 1 -->
</hh:paraPr>
```

- `type="OUTLINE"` + `level` 값으로 제목 계층 판별
- level 0 → `# 제목1`, level 1 → `## 제목2`, ...

### 스타일 매핑 체인

```
hp:p[paraPrIDRef] → hh:paraPr[id] → hh:heading[level]  → 제목 레벨
hp:run[charPrIDRef] → hh:charPr[id] → bold/italic/...   → 인라인 서식
```

## 5. 변환 파이프라인 작동 원리

### hwpx → markdown (`hwpx_to_md.py`)

```
hwpx (ZIP) → zipfile 해제
  ├── header.xml → lxml 파싱 → HwpxStyleMap 구축
  │   ├── paraPrIDRef → 제목 레벨 매핑
  │   └── charPrIDRef → 굵기/기울임/취소선 매핑
  ├── section*.xml → lxml 파싱 → 트리 순회 (다중 섹션 지원)
  │   ├── hp:p → 문단 → 마크다운 텍스트/제목
  │   ├── hp:tbl → 표 → 마크다운 테이블
  │   ├── hp:pic → 이미지 → ![ref](images/...)
  │   ├── hp:run/hp:t → 인라인 서식 적용
  │   ├── hp:header/hp:footer → <!-- 머리글/꼬리글 -->
  │   ├── hp:equation → $$ 수식 $$
  │   ├── hp:rect/hp:drawText → > [글상자] text
  │   ├── hp:checkBtn/radioBtn/comboBox/btn/edit → 양식 개체
  │   ├── hp:footnote/hp:endnote → [^N] 인라인 참조
  │   ├── hp:dutmal → <ruby>본말<rt>닷말</rt></ruby>
  │   ├── hp:textart → > [글맵시] text
  │   ├── hp:ole → <!-- OLE: comment -->
  │   ├── hp:deleteBegin/End → ~~취소선~~
  │   ├── hp:fieldBegin HYPERLINK → [text](url)
  │   ├── hp:colPr → <!-- [N단 레이아웃] -->
  │   └── 섹션 간 --- 구분자
  ├── BinData/* → 이미지 추출 (BMP→PNG 변환)
  └── template_info.json → 페이지 설정/양식 메타 보존
```

### markdown → hwpx (`md_to_hwpx.py` → pypandoc-hwpx)

```
markdown → Pandoc → JSON AST → PandocToHwpx 변환기
  ├── JSON AST 블록 순회 (Header, Para, Table, BulletList, Image...)
  ├── header.xml에서 스타일 ID 매핑
  ├── 새 charPr 생성 (굵기, 기울임 조합별)
  ├── 표 XML 생성 (rowspan/colspan 포함)
  ├── 이미지 임베딩 (BinData/에 복사)
  └── reference-doc의 원본 hwpx에서:
      ├── header.xml (스타일 정의) 복사
      ├── section0.xml (페이지 설정) 추출
      └── BinData/* (기존 이미지) 유지
```

### 왕복 변환 (양식 보존)

```
원본.hwpx → hwpx_to_md.py → 문서.md + template_info.json + images/
                                ↓ (AI로 내용 편집)
편집된.md → md_to_hwpx.py --reference-doc=원본.hwpx → 최종.hwpx
```

`--reference-doc`으로 원본 hwpx를 지정하면:
- 페이지 크기/여백/방향 유지
- 폰트/스타일 정의 유지
- 기존 이미지 파일 유지

## 6. 샘플 문서 통계 (신청서.hwpx)

| 항목 | 수량 |
|------|------|
| 전체 문단 (hp:p) | 991개 |
| 최상위 문단 | 298개 |
| 표 (hp:tbl) | 46개 |
| 텍스트 런 (hp:run) | 1,459개 |
| 텍스트 노드 (hp:t) | 1,233개 |
| 이미지 | 12개 (BMP 11, JPG 1) |
| 페이지 크기 | 59528 x 84188 (A4 세로) |

## 7. 삭제된 hwpx-to-markdown-v0.2 (Gemini 기반 웹앱) 상세 기술 문서

원래 Google AI Studio에서 생성한 React 웹앱으로, Gemini API를 사용하여 HWPX XML을 마크다운으로 변환하던 도구.
현재는 lxml 기반 로컬 파서(`hwpx_to_md.py`)로 대체하여 삭제함.
원본 AI Studio: https://aistudio.google.com/apps/drive/1NWb9Oww-4c9-0GQICEvJVJv4XPx6xfET?showPreview=true&showAssistant=true

### 7.1 기술 스택

| 기술 | 용도 |
|------|------|
| React 18 + TypeScript | UI 프레임워크 |
| Vite | 빌드 도구 |
| Tailwind CSS | 스타일링 (CSS 변수 기반 테마) |
| Google Material Symbols | 아이콘 |
| `@google/genai` | Gemini API SDK |
| `jszip` | 브라우저에서 HWPX(ZIP) 해제 |

### 7.2 아키텍처 & 컴포넌트 구조

```
App.tsx                          # 상태 관리 (history, activeItem, currentFile)
├── Header.tsx                   # 상단바 (로고, 다크/라이트 테마 토글)
├── Sidebar.tsx                  # 변환 이력 목록
├── MainConverter.tsx            # 핵심: 파일 업로드 → ZIP 해제 → XML 청킹 → Gemini 스트리밍
│   ├── ProcessingState.tsx      # 변환 중 진행률 UI (원형 프로그레스, 3단계 스텝)
│   └── ResultView.tsx           # 결과 뷰 (마크다운 raw + 대시보드 전환)
│       └── DashboardView.tsx    # 구조화된 대시보드 시각화
├── UnsupportedFormatModal.tsx   # .hwp 파일 업로드 시 경고 모달
└── services/
    └── geminiService.ts         # Gemini API 호출 (변환 + 대시보드 생성)
```

### 7.3 데이터 흐름

```
[사용자]
  │
  ▼ .hwpx 파일 드래그&드롭 또는 클릭 업로드
[App.tsx] ─── handleFileUpload()
  │  ├─ .hwp → UnsupportedFormatModal 표시
  │  ├─ .hwpx/.md → HistoryItem 생성, activeItem 설정
  │  └─ 그 외 → alert("지원하지 않는 파일 형식")
  │
  ▼ "고속 변환 시작" 버튼 클릭
[MainConverter.tsx] ─── startConversion()
  │
  ├─ Step 1: 아카이브 구조 분석
  │  └─ JSZip.loadAsync(file) → ZIP 해제
  │
  ├─ Step 2: 텍스트 레이어 추출
  │  └─ Contents/section*.xml 파일 추출 → 합본 XML 문자열
  │
  └─ Step 3: 무결성 스트리밍 변환
     ├─ XML을 40KB 청크로 분할
     ├─ 각 청크 → convertXmlChunkToMarkdownStream() (Gemini API)
     ├─ 스트리밍 응답을 실시간으로 UI에 렌더링
     └─ 청크 간 100ms 딜레이 (429 Rate Limit 방지)
  │
  ▼ 변환 완료
[ResultView.tsx]
  ├─ "무결성 문서" 탭: 마크다운 원문 표시 (pre 태그, 실시간 스트리밍 커서)
  ├─ "전략 대시보드" 탭: generateVisualDashboard() → DashboardView 렌더링
  ├─ 마크다운 복사 (clipboard API)
  └─ 다운로드 (마크다운 .md 또는 HTML 보고서)
```

### 7.4 Gemini API 서비스 상세 (`geminiService.ts`)

#### 7.4.1 XML → Markdown 스트리밍 변환

```typescript
// 핵심 함수: async generator로 스트리밍 텍스트 yield
async function* convertXmlChunkToMarkdownStream(xmlChunk: string)

// 모델: gemini-3-flash-preview
// temperature: 0.0 (결정론적 출력)
// 방식: generateContentStream (스트리밍)
```

**시스템 프롬프트 (Hallucination Zero 패턴)**:
```
[역할: 초고속 무결성 한글 문서 데이터 변환기]
당신은 창작자가 아닙니다. 입력된 HWPX XML 데이터에서 <hp:t> 태그 내의 텍스트를
추출하여 마크다운으로 변환하십시오.

[엄격 지침]
1. Hallucination ZERO: 원본에 없는 단어, 조사, 숫자를 절대 추가하지 마십시오.
2. No Commentary: "알겠습니다" 등 사족을 절대 붙이지 마십시오. 오직 본문만 출력.
3. Table & Structure: 표는 마크다운 표 형식을 엄격히 유지.
4. Speed: 텍스트 레이어를 즉시 추출하여 스트리밍.
```

**후처리**: 응답에서 ` ```markdown ``` ` 코드 블록 마커를 정규식으로 제거.

**에러 핸들링**: 429 에러(할당량 초과) 별도 메시지, 그 외 일반 통신 오류 메시지.

#### 7.4.2 대시보드 시각화 데이터 생성

```typescript
async function generateVisualDashboard(markdown: string): Promise<DashboardData>

// 모델: gemini-3-flash-preview
// temperature: 0.0
// 입력: markdown.slice(0, 30000) — 최대 30,000자 제한
// 출력: JSON (responseMimeType: "application/json")
// 스키마 검증: responseSchema로 구조 강제
```

**DashboardData 스키마**:
```typescript
interface DashboardData {
  title: string;                    // 문서 제목
  subtitle?: string;                // 부제목
  summary: string;                  // 핵심 요약
  stats: {                          // 주요 지표 (4개 카드 그리드)
    label: string;                  // 지표명
    value: string;                  // 값
    icon: string;                   // Material Symbol 아이콘명
    color?: string;
  }[];
  criticalNotice?: {                // 중요 알림 배너 (선택)
    title: string;
    content: string;
    icon: string;
  };
  sections: DashboardSection[];     // 본문 섹션 카드들
}

interface DashboardSection {
  id: string;
  type: 'card' | 'process' | 'table' | 'grid';  // 렌더링 방식
  title: string;
  icon: string;                     // Material Symbol 아이콘명
  content: string | any[];          // 텍스트 또는 표 데이터
  columns?: string[];               // 표일 때 컬럼 헤더
}
```

### 7.5 XML 청킹 전략 (`MainConverter.tsx`)

```
전체 section*.xml 합본 → 40,000자(~40KB) 단위로 분할

청크1 ──→ Gemini API (스트리밍) ──→ 실시간 UI 렌더
  ↓ 100ms 대기
청크2 ──→ Gemini API (스트리밍) ──→ 실시간 UI 렌더
  ↓ 100ms 대기
...
```

- `Contents/section*.xml` 파일을 숫자순으로 정렬 후 합본
- 40KB 고정 크기 슬라이싱 (XML 태그 경계 무시 — AI가 불완전 XML 처리)
- 청크 간 100ms `setTimeout` 딜레이로 429 Rate Limit 완화
- 진행률: `(처리된 청크 수 / 전체 청크 수) × 100` (최대 99%, 완료 시 100%)

### 7.6 UI/UX 디자인 패턴

#### 테마 시스템
- CSS 변수 기반 다크/라이트 모드: `--theme-*` 변수군
- `Header.tsx`에서 `document.documentElement.classList`로 `.dark`/`.light` 클래스 토글
- 모든 컴포넌트가 `var(--theme-*)` 참조 → 일괄 전환

#### 프로세싱 UI (`ProcessingState.tsx`)
- 3단계 스텝 인디케이터: 아카이브 분석 → 텍스트 추출 → 무결성 스트리밍
- SVG 원형 프로그레스바 (strokeDasharray/strokeDashoffset 애니메이션)
- 진행 중인 스텝: 회전 sync 아이콘 + "In Progress" 펄스 텍스트
- 완료된 스텝: done_all 체크 아이콘 + opacity 감소

#### 결과 뷰 (`ResultView.tsx`)
- 듀얼 뷰 탭: "무결성 문서" (마크다운 원문) / "전략 대시보드" (구조화 시각화)
- 스트리밍 중 자동 스크롤 + 깜빡이는 커서 애니메이션
- 복사: `navigator.clipboard.writeText()`
- 다운로드: 마크다운 모드 → `.md` 파일, 대시보드 모드 → 스타일링된 HTML 보고서

#### 대시보드 뷰 (`DashboardView.tsx`)
- 히어로 헤더: 문서 제목 (최대 9xl 크기), 부제목, 요약
- 통계 그리드: 4열 카드 레이아웃, 호버 시 아이콘 색상 반전
- 중요 알림: 주황색 경고 배너 (pulse 애니메이션)
- 섹션 카드: 12열 그리드, `card`/`process`는 6열, `table`/`grid`는 12열

#### HTML 보고서 내보내기 (`ResultView.tsx`)
- 대시보드 모드에서 다운로드 시 자체 스타일링된 HTML 문서 생성
- 인라인 CSS (외부 의존성 없음), 카드 레이아웃, 통계 그리드
- 파일명: `Report_{문서제목}.html`

### 7.7 삭제 이유 및 한계점

| 문제점 | 설명 |
|--------|------|
| API 의존성 | Gemini API 키 필수, 네트워크 연결 필수 |
| 비용 | 매 변환마다 API 호출 비용 발생 |
| 429 Rate Limit | 대용량 문서 시 할당량 초과 빈발 |
| XML 청킹 정확도 | 40KB 고정 슬라이싱으로 XML 태그가 잘림 → AI 해석에 의존 |
| 단방향 변환 | hwpx→md만 가능, md→hwpx 변환 없음 |
| Hallucination 위험 | temperature=0.0이라도 AI 특성상 100% 보장 불가 |

### 7.8 향후 재활용 가능한 아이디어

1. **DashboardData 시각화 시스템**
   - 변환된 마크다운을 구조화된 대시보드로 자동 요약
   - 사업계획서 검토/발표용 요약 뷰에 활용 가능
   - `responseSchema`로 JSON 구조를 강제하는 패턴 재사용 가능

2. **Hallucination Zero 프롬프트 패턴**
   - "당신은 창작자가 아닙니다" + temperature=0.0 조합
   - 원본 무결성이 중요한 모든 변환 작업에 적용 가능

3. **스트리밍 + 실시간 렌더링 UX**
   - `async generator` + `for await` 패턴으로 실시간 텍스트 표시
   - 프로그레스바와 스텝 인디케이터 조합

4. **HTML 보고서 내보내기**
   - 인라인 CSS만으로 독립 실행 가능한 보고서 생성
   - 외부 의존성 없이 이메일/공유에 적합

## 8. pypandoc-hwpx 분석 (GitHub: msjang/pypandoc-hwpx)

### 8.1 프로젝트 개요

| 항목 | 내용 |
|------|------|
| 저장소 | https://github.com/msjang/pypandoc-hwpx |
| 버전 | 0.1.1 (2025-12-18 릴리즈) |
| 라이선스 | MIT |
| Stars | 131 |
| Issues | 0개 (열린/닫힌 모두 없음) |
| 의존성 | Python 3.6+, Pandoc, pypandoc |

### 8.2 핵심 동작 원리

```
입력 파일 (md/docx/html)
  → Pandoc → JSON AST
    → PandocToHwpx.py → HWPX (ZIP)
      ├── header.xml (스타일: reference-doc에서 복사)
      ├── section0.xml (본문: AST에서 생성)
      ├── BinData/* (이미지 임베딩)
      └── 기타 메타데이터
```

### 8.3 `--reference-doc` 동작 방식

reference-doc(원본 hwpx)을 지정하면:
1. `header.xml` 전체를 원본에서 복사 (스타일 정의 보존)
2. `section0.xml`에서 `<hs:sec>` ~ 첫 번째 `<hp:p>` 사이의 **secPr(페이지 설정)**을 추출
3. 새로 생성된 본문 XML 앞에 원본 secPr을 접두어로 붙임
4. `_ensure_table_border_fill()`로 테이블용 borderFill이 없으면 새로 생성

**미지정 시**: 패키지 내장 `blank.hwpx`를 템플릿으로 사용 (Mac Word 16.73 기반, borderFill 2개)

### 8.4 발견된 버그 (pypandoc-hwpx 0.1.1)

#### 버그 1: borderFillIDRef 하드코딩 (PandocToHwpx.py:721)

```python
# 문제 코드 (PandocToHwpx.py 721행)
tbl_xml = f'<hp:tbl ... borderFillIDRef="3" ...'
```

- `hp:tbl`의 `borderFillIDRef`가 항상 `"3"`으로 하드코딩
- `blank.hwpx` 사용 시에는 `_ensure_table_border_fill()`이 ID 3의 borderFill을 생성하므로 문제없음
- **`--reference-doc` 사용 시**: 원본 header.xml에서 ID 3이 테이블 보더가 아닌 다른 용도일 수 있음
- **결과**: 테이블 테두리가 의도치 않은 스타일로 표시됨

#### 버그 2: 빈 마크다운 셀 → 빈 hp:subList (한글 크래시)

```
마크다운:  |  |  ← 빈 셀
생성 XML:  <hp:subList listType="PARA"></hp:subList>  ← hp:p 없음!
```

- 빈 마크다운 셀(`|  |`)을 변환할 때 `hp:subList` 안에 `hp:p` 자식 요소를 넣지 않음
- **한컴오피스 한글이 빈 subList에서 크래시** (프로그램 강제 종료)
- 이것이 roundtrip hwpx가 열리지 않는 **주요 원인**

#### 우리의 해결책 (`md_to_hwpx.py._patch_hwpx()`)

```python
# 패치 1: 빈 subList에 빈 paragraph 주입
empty_para = '<hp:p paraPrIDRef="..." ...><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run></hp:p>'

# 패치 2: hp:tbl의 borderFillIDRef를 첫 번째 셀의 값으로 통일
```

> GitHub Issues에 이 버그들이 보고되어 있지 않음 (0개 이슈). 향후 이슈 제출 고려.

### 8.5 Pandoc JSON AST의 근본적 한계

pypandoc-hwpx는 Pandoc의 JSON AST를 입력으로 받기 때문에, AST에 포함되지 않는 정보는 변환 불가:

| 손실되는 정보 | 설명 |
|--------------|------|
| 셀 너비 (cellSz) | 마크다운에는 열 너비 개념 없음 → 균등 분배 |
| 셀 병합 (cellSpan) | 마크다운 표는 colspan/rowspan 미지원 |
| 표 안의 표 | Pandoc AST가 중첩 테이블 미지원 |
| 셀 배경색 | 마크다운에 해당 문법 없음 |
| 장평/자간 | Pandoc AST에 미포함 |
| 페이지 나누기 위치 | AST에 미포함 |

## 9. 원본 vs 변환후 비교 분석 (스크린샷)

### 9.1 원본 문서 특징 (원본.png)

- **복잡한 표 구조**: 46개 테이블, 다양한 셀 병합 (2열 병합, 3행 병합 등)
- **정교한 셀 크기**: 열마다 다른 너비, 행마다 다른 높이
- **서식 풍부**: 굵은 헤더, 셀 배경색 (회색/파란색/주황색), 테두리 스타일 차이
- **이미지 삽입**: 표 안에 이미지 배치, 정확한 크기 지정
- **표 안의 표**: 한글 특유의 중첩 표 구조
- **다단 레이아웃**: 좌우 구분된 셀 레이아웃

### 9.2 변환 후 문제점 (변환후.png — pypandoc-hwpx roundtrip)

| 문제 | 심각도 | 원인 |
|------|--------|------|
| 셀 병합 해제 — 병합되었던 셀이 개별 셀로 분리 | 치명적 | 마크다운에 colspan/rowspan 문법 없음 |
| 균등 열 너비 — 모든 열이 동일한 너비 | 치명적 | pypandoc-hwpx가 열 너비를 균등 분배 |
| 표 안의 표 소실 — 중첩 표가 평면화 | 치명적 | Pandoc AST 미지원 |
| 셀 배경색 소실 — 모든 셀 배경이 흰색 | 중요 | 마크다운에 셀 색상 문법 없음 |
| 행 높이 균등화 — 원래 다양했던 높이가 통일 | 중요 | pypandoc-hwpx가 높이 균등 분배 |
| 이미지 위치 변경 — 표 안 이미지가 표 밖으로 | 중요 | 마크다운 표 안 이미지 한계 |
| 텍스트 서식 부분 손실 — 셀 내 굵기/색상 차이 소실 | 경미 | 일부 인라인 서식 보존됨 |
| 빈 행/열 추가 — 원본에 없는 빈 셀 생성 | 경미 | 병합 해제로 인한 부산물 |

### 9.3 스마트 교체(smart_replace.py)로 해결되는 문제

| 문제 | pypandoc roundtrip | 스마트 교체 |
|------|-------------------|-------------|
| 셀 병합 | 해제됨 | **보존** |
| 셀 너비/높이 | 균등화 | **보존** |
| 표 안의 표 | 소실 | **보존** |
| 셀 배경색 | 소실 | **보존** |
| 테두리 스타일 | 변경됨 | **보존** |
| 이미지 위치 | 이동됨 | **보존** |
| 텍스트 내용 변경 | 반영 | **반영** |
| 새 표/행/열 추가 | 가능 | 불가 |
| 문단 추가/삭제 | 가능 | 불가 |

## 10. 스마트 교체 (`smart_replace.py`)

### 10.1 개요

pypandoc-hwpx의 표 서식 손실 문제를 근본적으로 해결하기 위한 대안 변환 방식.
원본 HWPX의 XML 구조를 그대로 유지하면서, 편집된 마크다운의 **텍스트 내용만** 반영.

### 10.2 동작 원리

```
원본.hwpx → section*.xml 파싱 → XML 블록 추출 (다중 섹션 지원)
편집된.md → 마크다운 파싱  → MD 블록 추출
                ↓
        위치 기반 매칭 (OWPML 2024 네임스페이스 자동 감지)
                ↓
    XML의 hp:t 텍스트만 교체 (동적 네임스페이스 접두사 사용)
                ↓
        원본 HWPX + 수정된 section*.xml → 출력.hwpx
```

**주요 기능**:
- 다중 섹션 처리 (`section0~N.xml` 전체 처리)
- OWPML 2024 네임스페이스 자동 감지 (`NS_2011`/`NS_2024` 매핑)
- 동적 네임스페이스 접두사 감지 (실제 XML의 접두사 사용)

### 10.3 사용법

```bash
# 단독 실행
python smart_replace.py 원본.hwpx 편집된.md -o 최종본.hwpx

# convert.py 통합 CLI
python convert.py smart 원본.hwpx 편집된.md -o 최종본.hwpx
```

### 10.4 권장 워크플로

```
1. python convert.py to-md 원본.hwpx -o output/문서.md
2. AI로 output/문서.md 편집 (텍스트 내용만 변경)
3. python convert.py smart 원본.hwpx output/문서.md -o 최종본.hwpx
```

**주의**: 마크다운에서 표 행/열을 추가/삭제하면 매칭이 어긋남.
텍스트 내용 편집만 하고, 구조 변경이 필요하면 원본 hwpx를 직접 수정할 것.

## 11. 의존성

| 패키지 | 버전 | 용도 |
|--------|------|------|
| Python | 3.13.2 | 런타임 |
| Pandoc | 3.9 | pypandoc-hwpx 내부 사용 |
| pypandoc-hwpx | 0.1.1 | md → hwpx 변환 (to-hwpx 모드) |
| pypandoc | 1.16.2 | Pandoc Python 래퍼 |
| lxml | 6.0.2 | HWPX XML 파싱 (smart 모드 + to-md) |
| Pillow | 12.0.0 | BMP → PNG 이미지 변환 |

## 12. 기능 커버리지 분석 (hwpxlib 테스트 파일 기반)

hwpxlib 프로젝트(`github.com/neolord0/hwpxlib`)의 47개 테스트 HWPX 파일로 파이프라인을 검증했다.

### 12.1 기능별 테스트 결과

| 테스트 파일 | 기능 | XML 태그 | 결과 | 비고 |
|---|---|---|---|---|
| SimpleTable.hwpx | 표 | `hp:tbl/hp:tr/hp:tc` | ✅ 정상 | 4줄 60자 추출 |
| sample1.hwpx | 텍스트+서식 | `hp:run/hp:t` | ✅ 정상 | *수학* 이탤릭 |
| SimplePicture.hwpx | 이미지 | `hp:pic` | ✅ 정상 | `![image1]()` |
| MultiColumn.hwpx | 다단 | `hp:colPr colCount` | ✅ P3 구현 | `<!-- [3단 레이아웃] -->` |
| ChangeTrack.hwpx | 변경 추적 | `hp:deleteBegin/End` | ✅ P3 구현 | `~~삭제~~ 삽입` |
| SimpleDutmal.hwpx | 덧말(ruby) | `hp:dutmal` | ✅ P3 구현 | `<ruby>본말<rt>닷말</rt></ruby>` |
| SimpleTextArt.hwpx | 글맵시 | `hp:textart` | ✅ P3 구현 | `> [글맵시] text` |
| SimpleOLE.hwpx | OLE 개체 | `hp:ole` | ✅ P3 구현 | `<!-- OLE: comment -->` |
| SimpleEquation.hwpx | 수식 | `hp:equation/hp:script` | ✅ P1 구현 | `$$ script $$` |
| HeaderFooter.hwpx | 머리글/꼬리글 | `hp:header/hp:footer` | ✅ P1 구현 | `<!-- 머리글: text -->` |
| SimpleComboBox.hwpx | 콤보박스 | `hp:comboBox` | ✅ P2 구현 | `[콤보: name]` |
| SimpleButtons.hwpx | 버튼/체크/라디오 | `hp:btn/hp:checkBtn/hp:radioBtn` | ✅ P2 구현 | `[ ] caption` |
| SimpleRectangle.hwpx | 도형(글상자) | `hp:rect/hp:drawText` | ✅ P1 구현 | `> [글상자] ABC` |
| SimpleEdit.hwpx | 입력란 | `hp:edit` | ✅ P2 구현 | `[입력란: name]` |
| 3-section test.hwpx | 다중 섹션 | `section0~2.xml` | ✅ P2 구현 | `---` 구분자 |
| 재난안전종합상황.hwpx | 대용량(7.3MB) | 복합 | ✅ 처리 | 1599줄 65KB |
| **전체 25개** | **reader_writer 일괄** | — | ✅ **에러 0** | — |

### 12.2 미지원 기능의 XML 구조

**수식** — `hp:equation` 내 `hp:script`에 한글 수식 문법 저장:
```xml
<hp:equation ...><hp:script>{"123"} over {123 sqrt {3466}}</hp:script></hp:equation>
```

**머리글/꼬리글** — `hp:header/hp:footer` 내 `hp:subList` → `hp:p` 구조:
```xml
<hp:header applyPageType="BOTH"><hp:subList ...>
  <hp:p ...><hp:run ...><hp:t>머리말 테스트</hp:t></hp:run></hp:p>
</hp:subList></hp:header>
```

**양식 개체** — `hp:checkBtn/hp:radioBtn/hp:comboBox/hp:btn` 속성에 텍스트:
```xml
<hp:checkBtn caption="선택 상자1" value="UNCHECKED" name="CheckBox1" .../>
<hp:radioBtn caption="라디오 단추1" value="UNCHECKED" name="RadioButton1" .../>
<hp:comboBox name="ComboBox1" selectedValue="" .../>
```

**도형 안 텍스트** — `hp:rect` → `hp:drawText` → `hp:subList` → `hp:p`:
```xml
<hp:rect ...><hp:drawText ...><hp:subList ...>
  <hp:p ...><hp:run ...><hp:t>ABC</hp:t></hp:run></hp:p>
</hp:subList></hp:drawText></hp:rect>
```

### 12.3 개선 우선순위

**P1 — 텍스트 손실 방지 ✅ 완료 (커밋 1e16b9f)**
1. ✅ 도형/글상자 안 텍스트 → `> [글상자] text`
2. ✅ 머리글/꼬리글 → `<!-- 머리글: text -->`
3. ✅ 수식 → `$$ script $$`

**P2 — 양식 문서 + 구조 확장 ✅ 완료 (커밋 caa49e8)**
4. ✅ 체크박스/라디오 → `[ ] caption` / `[x] caption` / `(o) caption`
5. ✅ 콤보박스/입력란/버튼 → `[콤보: name]` / `[입력란: name]` / `[버튼: caption]`
6. ✅ 각주/미주 → 인라인 `[^N]` + 문서 끝 `[^N]: text`
7. ✅ 다중 섹션 → section*.xml 자동 탐색, `---` 구분자

**P3 — 잔여 기능 + 텍스트 완전성 ✅ 완료 (커밋 30d2551)**
8. ✅ 덧말(ruby text) → `<ruby>본말<rt>닷말</rt></ruby>`
9. ✅ TextArt(글맵시) → `> [글맵시] text` (text 속성에서 추출)
10. ✅ OLE 개체 → `<!-- OLE: comment (ref) -->` 주석
11. ✅ 변경 추적 → `~~삭제~~` 취소선 + 삽입 그대로
12. ✅ 하이퍼링크 → `[text](url)` (방어적 구현)
13. ✅ 다단 레이아웃 → `<!-- [N단 레이아웃] -->` 메타데이터

**P4 — OWPML 2024 호환 ✅ 완료 (커밋 8f47702)**
14. ✅ 네임스페이스 자동 감지 (`hancom.co.kr/2011` ↔ `owpml.org/2024`)
15. ✅ NS_2011/NS_2024 매핑 테이블 분리, section→body 변경 대응
16. ✅ smart_replace.py 2024 네임스페이스 + 다중 섹션 지원
17. ✅ md_to_hwpx.py 다중 섹션 패치 지원

**P5 — 잔여 기능 + 품질 ✅ 완료 (커밋 fa994ad)**
18. ✅ 책갈피 (`hp:bookmarkStart`) → `<a id="name"></a>` 앵커
19. ✅ 탭 (`hp:tab`) → 공백 4칸 변환
20. ✅ smart_replace.py 일반 문단 텍스트 교체 지원
21. ✅ pytest 테스트 스위트 (14 passed, 1 skipped)
22. ✅ README.md 사용법 가이드 문서

### 12.4 smart_replace.py v2 개선 사항

프래그먼트 레벨 diff 폴백 추가 (2025-02):
- 전략 1: 전체 셀 텍스트 매칭 (단일 `hp:t` 셀)
- 전략 2: `difflib.SequenceMatcher`로 변경 단어만 교체 (멀티런 셀)
- `_normalize()`에서 `*` 제거 (마크다운 라운드트립 아티팩트 방지)

v3 개선 (2025-02):
- OWPML 2024 네임스페이스 자동 감지 (`NS_2011`/`NS_2024` 매핑)
- 다중 섹션 지원 (`section0~N.xml` 전체 처리)
- 동적 네임스페이스 접두사 감지 (`</hp:t>` 대신 실제 접두사 사용)

v4 개선 (2025-02):
- 일반 문단 텍스트 교체 지원 (`extract_xml_paragraphs()` + `parse_markdown_paragraphs()`)
- 테이블 + 문단 동시 교체로 더 완전한 라운드트립

### 12.5 테스트 데이터

hwpxlib 테스트 파일 (`test_data/hwpxlib/testFile/`):
- `reader_writer/` — 기능별 단위 테스트 (25개)
- `error/` — 실제 문서 에러 케이스 (15개)
- `tool/` — 텍스트 추출기 테스트 (7개)
