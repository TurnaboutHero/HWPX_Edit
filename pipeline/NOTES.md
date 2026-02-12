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
  ├── section0.xml → lxml 파싱 → 트리 순회
  │   ├── hp:p → 문단 → 마크다운 텍스트/제목
  │   ├── hp:tbl → 표 → 마크다운 테이블
  │   ├── hp:pic → 이미지 → ![ref](images/...)
  │   └── hp:run/hp:t → 인라인 서식 적용
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

## 8. 의존성

| 패키지 | 버전 | 용도 |
|--------|------|------|
| Python | 3.13.2 | 런타임 |
| Pandoc | 3.9 | pypandoc-hwpx 내부 사용 |
| pypandoc-hwpx | 0.1.1 | md → hwpx 변환 |
| pypandoc | 1.16.2 | Pandoc Python 래퍼 |
| lxml | 6.0.2 | HWPX XML 파싱 |
| Pillow | 12.0.0 | BMP → PNG 이미지 변환 |
