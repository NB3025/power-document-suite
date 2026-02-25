# PowerPoint 프레젠테이션 생성 (PptxGenJS)

PptxGenJS를 사용한 네이티브 PowerPoint 객체 기반 프레젠테이션 생성 가이드입니다.

## Overview

새로운 PowerPoint 프레젠테이션을 처음부터 만들 때 사용합니다. 모든 텍스트, 도형, 차트, 테이블이 PowerPoint에서 편집 가능한 네이티브 객체로 생성됩니다.

## Workflow Decision Tree

```
새 프레젠테이션 생성?
├── 예 → 이 가이드 (PptxGenJS)
└── 아니오 (기존 템플릿 수정) → pptx-template.md
```

## Prerequisites

```bash
# 필수
npm install -g pptxgenjs

# 아이콘 사용 시
npm install -g react react-dom react-icons sharp

# QA용
pip install "markitdown[pptx]" Pillow
brew install poppler  # macOS (pdftoppm)
```

---

## 워크플로우

```
1. 디자인 결정 (색상 팔레트, 폰트, 비주얼 모티프)
   ↓
2. PptxGenJS로 슬라이드 작성 (네이티브 텍스트/도형/차트)
   ↓
3. QA: 콘텐츠 확인 (markitdown)
   ↓
4. QA: 시각 검증 (soffice → pdftoppm → 이미지 검토)
   ↓
5. 문제 수정 → 재검증 반복
```

---

## Design Principles

**지루한 슬라이드를 만들지 마세요.** 흰 배경에 글머리 기호만 있는 슬라이드는 아무도 감동시키지 못합니다.

### 시작 전 결정사항

- **주제에 맞는 대담한 색상 팔레트 선택**: 아무 프레젠테이션에나 넣어도 어울리는 색상이면 충분히 구체적이지 않은 것
- **색상 비중 차등화**: 하나의 색이 60-70% 차지, 1-2개 보조색, 하나의 강조색. 모든 색에 동일 비중 금지
- **명암 대비**: 타이틀/결론은 어두운 배경, 콘텐츠는 밝은 배경 ("샌드위치" 구조). 또는 전체 다크 테마
- **비주얼 모티프 통일**: 하나의 특징적 요소를 선택하고 반복 (둥근 이미지 프레임, 색상 원 안 아이콘, 두꺼운 한쪽 테두리 등)

### 색상 팔레트

주제에 맞는 색상을 선택하세요. 기본 파란색으로 기본값 설정하지 마세요.

| 테마 | Primary | Secondary | Accent |
|------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

### 폰트 페어링

기본 Arial만 쓰지 마세요. 헤더와 본문을 구분하세요.

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Calibri | Calibri Light |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Impact | Arial |
| Palatino | Garamond |
| Consolas | Calibri |

| 요소 | 크기 |
|------|------|
| 슬라이드 제목 | 36-44pt bold |
| 섹션 헤더 | 20-24pt bold |
| 본문 텍스트 | 14-16pt |
| 캡션 | 10-12pt muted |

### 각 슬라이드별 레이아웃 아이디어

**모든 슬라이드에 비주얼 요소 필수** (이미지, 차트, 아이콘, 도형). 텍스트만 있는 슬라이드는 금지.

**레이아웃 옵션:**
- 2단 (텍스트 왼쪽, 일러스트 오른쪽)
- 아이콘 + 텍스트 행 (색상 원 안 아이콘, 볼드 헤더, 설명)
- 2x2 또는 2x3 그리드 (한쪽 이미지, 다른쪽 콘텐츠 블록)
- 반 블리드 이미지 (전체 좌측 또는 우측) + 콘텐츠 오버레이

**데이터 표시:**
- 큰 통계 콜아웃 (60-72pt 큰 숫자 + 아래 작은 라벨)
- 비교 컬럼 (전/후, 장/단점, 나란히 옵션)
- 타임라인 또는 프로세스 플로우 (번호 단계, 화살표)

### 간격

- 최소 0.5" 마진
- 콘텐츠 블록 사이 0.3-0.5"
- 여백을 충분히 남기기

### 피해야 할 것 (Common Mistakes)

- **같은 레이아웃 반복 금지** — 슬라이드마다 컬럼, 카드, 콜아웃을 다양하게
- **본문 텍스트 가운데 정렬 금지** — 단락과 리스트는 왼쪽 정렬, 제목만 가운데
- **크기 대비 부족 금지** — 제목 36pt+ vs 본문 14-16pt
- **기본 파란색 남용 금지** — 주제에 맞는 색상 선택
- **일부만 스타일링 금지** — 전체를 일관되게 스타일링하거나 심플하게 통일
- **텍스트만 있는 슬라이드 금지** — 이미지, 아이콘, 차트, 비주얼 요소 추가
- **제목 아래 악센트 라인 금지** — AI가 만든 느낌의 대표적 특징; 여백이나 배경색으로 대체
- **저대비 요소 금지** — 아이콘과 텍스트 모두 배경과 강한 대비 필요

---

## Setup & Basic Structure

```javascript
const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // 10" x 5.625"
pres.author = 'Author Name';
pres.title = 'Presentation Title';

let slide = pres.addSlide();
slide.addText("Hello World!", { x: 0.5, y: 0.5, fontSize: 36, color: "363636" });

pres.writeFile({ fileName: "output.pptx" });
```

### Layout Dimensions (좌표 단위: 인치)

- `LAYOUT_16x9`: 10" x 5.625" (기본)
- `LAYOUT_16x10`: 10" x 6.25"
- `LAYOUT_4x3`: 10" x 7.5"
- `LAYOUT_WIDE`: 13.3" x 7.5"

---

## 텍스트 & 포맷팅

```javascript
// 기본 텍스트
slide.addText("Simple Text", {
  x: 1, y: 1, w: 8, h: 2, fontSize: 24, fontFace: "Arial",
  color: "363636", bold: true, align: "center", valign: "middle"
});

// 자간 (charSpacing 사용, letterSpacing 금지 — 무시됨)
slide.addText("SPACED TEXT", { x: 1, y: 1, w: 8, h: 1, charSpacing: 6 });

// 리치 텍스트 배열
slide.addText([
  { text: "Bold ", options: { bold: true } },
  { text: "Italic ", options: { italic: true } }
], { x: 1, y: 3, w: 8, h: 1 });

// 여러 줄 텍스트 (breakLine: true 필수)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }  // 마지막 항목은 breakLine 불필요
], { x: 0.5, y: 0.5, w: 8, h: 2 });

// 텍스트 박스 마진 (내부 패딩)
slide.addText("Title", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  margin: 0  // 도형/아이콘과 정렬할 때 0으로 설정
});
```

**Tip:** 텍스트 박스는 기본 마진이 있음. 도형, 선, 아이콘과 정밀 정렬 시 `margin: 0` 설정.

---

## 리스트 & Bullet

```javascript
// ✅ CORRECT: 올바른 bullet 사용
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ WRONG: 유니코드 bullet 금지 (이중 bullet 발생)
slide.addText("* First item", { ... });

// 하위 항목 및 번호 리스트
{ text: "Sub-item", options: { bullet: true, indentLevel: 1 } }
{ text: "First", options: { bullet: { type: "number" }, breakLine: true } }
```

**⚠️ CRITICAL**: `lineSpacing`과 bullet을 함께 사용하지 마세요 — 과도한 간격 발생. `paraSpaceAfter` 사용.

---

## 도형

```javascript
// 사각형
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" }, line: { color: "000000", width: 2 }
});

// 원
slide.addShape(pres.shapes.OVAL, { x: 4, y: 1, w: 2, h: 2, fill: { color: "0000FF" } });

// 선
slide.addShape(pres.shapes.LINE, {
  x: 1, y: 3, w: 5, h: 0, line: { color: "FF0000", width: 3, dashType: "dash" }
});

// 투명도
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});

// 둥근 사각형 (ROUNDED_RECTANGLE에만 rectRadius 작동, RECTANGLE에는 안됨)
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// 그림자
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

### Shadow 옵션

| 속성 | 타입 | 범위 | 비고 |
|------|------|------|------|
| `type` | string | `"outer"`, `"inner"` | |
| `color` | string | 6자리 hex (예: `"000000"`) | `#` 접두사 없이, 8자리 hex 금지 |
| `blur` | number | 0-100 pt | |
| `offset` | number | 0-200 pt | **반드시 양수** — 음수는 파일 손상 |
| `angle` | number | 0-359도 | 그림자 방향 (135 = 우하단, 270 = 상단) |
| `opacity` | number | 0.0-1.0 | 투명도는 여기서 설정, 색상에 인코딩 금지 |

**Note**: 그라데이션 fill은 네이티브 미지원. 그라데이션 이미지를 배경으로 사용.

---

## 이미지

```javascript
// 파일 경로
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// URL
slide.addImage({ path: "https://example.com/image.jpg", x: 1, y: 1, w: 5, h: 3 });

// Base64 (더 빠름)
slide.addImage({ data: "image/png;base64,iVBORw0KGgo...", x: 1, y: 1, w: 5, h: 3 });
```

### 이미지 옵션

```javascript
slide.addImage({
  path: "image.png",
  x: 1, y: 1, w: 5, h: 3,
  rotate: 45,              // 0-359도
  rounding: true,          // 원형 크롭
  transparency: 50,        // 0-100
  altText: "Description",  // 접근성
  hyperlink: { url: "https://example.com" }
});
```

### 이미지 크기 모드

```javascript
// Contain - 비율 유지하며 영역 안에 맞춤
{ sizing: { type: 'contain', w: 4, h: 3 } }

// Cover - 비율 유지하며 영역 채움 (잘릴 수 있음)
{ sizing: { type: 'cover', w: 4, h: 3 } }

// Crop - 특정 부분만 잘라내기
{ sizing: { type: 'crop', x: 0.5, y: 0.5, w: 2, h: 2 } }
```

### 종횡비 유지 크기 계산

```javascript
const origWidth = 1978, origHeight = 923, maxHeight = 3.0;
const calcWidth = maxHeight * (origWidth / origHeight);
const centerX = (10 - calcWidth) / 2;

slide.addImage({ path: "image.png", x: centerX, y: 1.2, w: calcWidth, h: maxHeight });
```

---

## 아이콘

react-icons로 SVG 아이콘을 생성하고 PNG로 래스터화하여 사용합니다.

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle, FaChartLine } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// 슬라이드에 추가
const iconData = await iconToBase64Png(FaCheckCircle, "#4472C4", 256);
slide.addImage({ data: iconData, x: 1, y: 1, w: 0.5, h: 0.5 });
```

**Note**: size 256 이상으로 설정해야 선명함. size는 래스터화 해상도이고, 슬라이드 표시 크기는 `w`/`h`로 설정.

### 아이콘 라이브러리

- `react-icons/fa` - Font Awesome
- `react-icons/md` - Material Design
- `react-icons/hi` - Heroicons
- `react-icons/bi` - Bootstrap Icons

---

## 슬라이드 배경

```javascript
// 단색
slide.background = { color: "F1F1F1" };

// 투명도
slide.background = { color: "FF3399", transparency: 50 };

// 이미지 (URL)
slide.background = { path: "https://example.com/bg.jpg" };

// 이미지 (Base64)
slide.background = { data: "image/png;base64,iVBORw0KGgo..." };
```

### 그라데이션 배경 (이미지로 생성)

```javascript
const sharp = require('sharp');

async function createGradientBackground(color1, color2) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="563">
    <defs>
      <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" style="stop-color:#${color1}"/>
        <stop offset="100%" style="stop-color:#${color2}"/>
      </linearGradient>
    </defs>
    <rect width="100%" height="100%" fill="url(#g)"/>
  </svg>`;
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// 사용
slide.background = { data: await createGradientBackground('1E2761', '2E4053') };
```

---

## 테이블

```javascript
// 기본 테이블
slide.addTable([
  ["Header 1", "Header 2"],
  ["Cell 1", "Cell 2"]
], {
  x: 1, y: 1, w: 8, h: 2,
  border: { pt: 1, color: "999999" }, fill: { color: "F1F1F1" }
});

// 스타일 + 병합 셀
let tableData = [
  [
    { text: "Header", options: { fill: { color: "6699CC" }, color: "FFFFFF", bold: true } },
    "Cell"
  ],
  [{ text: "Merged", options: { colspan: 2 } }]
];
slide.addTable(tableData, { x: 1, y: 3.5, w: 8, colW: [4, 4] });
```

---

## 차트

### 기본 차트

```javascript
// Bar chart
slide.addChart(pres.charts.BAR, [{
  name: "Sales", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4500, 5500, 6200, 7100]
}], {
  x: 0.5, y: 0.6, w: 6, h: 3, barDir: 'col',
  showTitle: true, title: 'Quarterly Sales'
});

// Line chart
slide.addChart(pres.charts.LINE, [{
  name: "Temp", labels: ["Jan", "Feb", "Mar"], values: [32, 35, 42]
}], { x: 0.5, y: 4, w: 6, h: 3, lineSize: 3, lineSmooth: true });

// Pie chart
slide.addChart(pres.charts.PIE, [{
  name: "Share", labels: ["A", "B", "Other"], values: [35, 45, 20]
}], { x: 7, y: 1, w: 5, h: 4, showPercent: true });
```

### 모던한 차트 스타일링

기본 차트는 구식. 다음 옵션으로 깔끔한 모던 스타일 적용:

```javascript
slide.addChart(pres.charts.BAR, chartData, {
  x: 0.5, y: 1, w: 9, h: 4, barDir: "col",

  // 팔레트 색상 매칭
  chartColors: ["0D9488", "14B8A6", "5EEAD4"],

  // 깔끔한 배경
  chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },

  // 축 라벨 색상
  catAxisLabelColor: "64748B",
  valAxisLabelColor: "64748B",

  // 그리드 라인 (값 축만, 카테고리 축 숨김)
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },

  // 데이터 라벨
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelColor: "1E293B",

  // 단일 시리즈면 범례 숨김
  showLegend: false,
});
```

**차트 스타일링 주요 옵션:**
- `chartColors: [...]` — 시리즈/세그먼트 hex 색상
- `chartArea: { fill, border, roundedCorners }` — 차트 배경
- `catGridLine/valGridLine: { color, style, size }` — 그리드 라인 (`style: "none"`으로 숨김)
- `lineSmooth: true` — 곡선 (선 차트)
- `legendPos: "r"` — 범례 위치: "b", "t", "l", "r", "tr"

### 차트 유형

```javascript
pres.charts.BAR      // 막대 차트
pres.charts.LINE     // 선 차트
pres.charts.PIE      // 파이 차트
pres.charts.DOUGHNUT // 도넛 차트
pres.charts.SCATTER  // 산점도
pres.charts.BUBBLE   // 버블 차트
pres.charts.RADAR    // 레이더 차트
```

---

## Slide Masters

```javascript
pres.defineSlideMaster({
  title: 'TITLE_SLIDE', background: { color: '283A5E' },
  objects: [{
    placeholder: { options: { name: 'title', type: 'title', x: 1, y: 2, w: 8, h: 2 } }
  }]
});

let titleSlide = pres.addSlide({ masterName: "TITLE_SLIDE" });
titleSlide.addText("My Title", { placeholder: "title" });
```

---

## ⚠️ Common Pitfalls (파일 손상/버그 유발)

1. **색상에 "#" 금지** — 파일 손상
   ```javascript
   color: "FF0000"      // ✅ CORRECT
   color: "#FF0000"     // ❌ WRONG
   ```

2. **8자리 hex 색상에 opacity 인코딩 금지** — 파일 손상. `opacity` 속성 사용
   ```javascript
   shadow: { color: "00000020" }                        // ❌ 파일 손상
   shadow: { color: "000000", opacity: 0.12 }           // ✅ CORRECT
   ```

3. **유니코드 bullet 금지** — `bullet: true` 사용

4. **줄 바꿈은 `breakLine: true`** — 배열 항목 사이에 필수

5. **bullet과 `lineSpacing` 함께 사용 금지** — `paraSpaceAfter` 사용

6. **pptxgen() 인스턴스 재사용 금지** — 프레젠테이션마다 새로 생성

7. **옵션 객체 재사용 금지** — PptxGenJS가 내부적으로 객체를 변환함. 공유하면 두 번째 호출에서 손상
   ```javascript
   // ❌ WRONG
   const shadow = { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 };
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });  // 이미 변환된 값

   // ✅ CORRECT: 팩토리 함수 사용
   const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 });
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
   ```

8. **ROUNDED_RECTANGLE에 악센트 테두리 겹치기 금지** — 사각 오버레이가 둥근 모서리를 덮지 못함. RECTANGLE 사용
   ```javascript
   // ❌ WRONG: 악센트 바가 둥근 모서리를 덮지 못함
   slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { ... });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });

   // ✅ CORRECT: RECTANGLE으로 통일
   slide.addShape(pres.shapes.RECTANGLE, { ... });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });
   ```

---

## QA (필수)

**문제가 있다고 가정하고, 찾아내는 것이 목표입니다.**

첫 번째 렌더링은 거의 항상 문제가 있음. QA를 확인 절차가 아닌 버그 찾기로 접근하세요.

### 콘텐츠 QA

```bash
python -m markitdown output.pptx
```

누락 콘텐츠, 오탈자, 순서 오류 확인.

### 시각 QA

슬라이드를 이미지로 변환 후 검토:

```bash
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

이미지 검토 시 확인사항:
- 겹치는 요소 (텍스트가 도형을 관통, 선이 글자를 관통)
- 텍스트 오버플로우 또는 잘림
- 장식 선이 한 줄용인데 제목이 두 줄로 줄바꿈됨
- 요소 간 간격 부족 (< 0.3" 간격) 또는 불균일
- 슬라이드 가장자리 마진 부족 (< 0.5")
- 저대비 텍스트/아이콘
- 텍스트 박스가 너무 좁아 과도한 줄바꿈

### 검증 루프

1. 슬라이드 생성 → 이미지 변환 → 검토
2. **발견된 문제 리스트**
3. 문제 수정
4. **수정된 슬라이드 재검증** — 하나의 수정이 다른 문제를 만들 수 있음
5. 전체 패스에서 새 문제 없을 때까지 반복

**최소 한 번의 수정-검증 사이클을 완료하기 전에 성공 선언 금지.**

---

## Quick Reference

| 작업 | 코드 |
|------|------|
| 슬라이드 추가 | `pres.addSlide()` |
| 텍스트 추가 | `slide.addText("text", {x, y, w, h})` |
| 도형 추가 | `slide.addShape(pres.shapes.RECTANGLE, {x, y, w, h})` |
| 이미지 추가 | `slide.addImage({path: "img.png", x, y, w, h})` |
| 차트 추가 | `slide.addChart(pres.charts.BAR, data, opts)` |
| 테이블 추가 | `slide.addTable(rows, opts)` |
| 배경 설정 | `slide.background = { color: "F1F1F1" }` |
| 저장 | `pres.writeFile({ fileName: "output.pptx" })` |

### 도형 유형

`RECTANGLE`, `OVAL`, `LINE`, `ROUNDED_RECTANGLE`

### 레이아웃 상수

```javascript
pres.layout = 'LAYOUT_16x9';   // 10" x 5.625"
pres.layout = 'LAYOUT_16x10';  // 10" x 6.25"
pres.layout = 'LAYOUT_4x3';    // 10" x 7.5"
pres.layout = 'LAYOUT_WIDE';   // 13.3" x 7.5"
```

---

## Troubleshooting

### "File corrupted" / PowerPoint에서 열 수 없음
- **원인**: 색상 코드에 `#` 포함 또는 8자리 hex 사용
- **해결**: `"FF0000"` 형식 사용, opacity는 별도 속성으로

### 텍스트가 겹치거나 잘림
- **원인**: 텍스트 박스 크기 부족 또는 margin 미설정
- **해결**: w/h 값 조정, `margin: 0` 설정 확인

### Bullet이 이중으로 표시
- **원인**: 유니코드 bullet 문자 사용
- **해결**: `bullet: true` 옵션만 사용

### 그라데이션이 표시 안됨
- **원인**: 네이티브 그라데이션 미지원
- **해결**: Sharp로 그라데이션 PNG 생성 후 배경 이미지로 사용

### 여러 도형의 그림자가 이상함
- **원인**: shadow 옵션 객체 재사용
- **해결**: 팩토리 함수로 매번 새 객체 생성
