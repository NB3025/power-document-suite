# PowerPoint 프레젠테이션 생성 (HTML→PPTX)

HTML 슬라이드를 PowerPoint로 변환하는 html2pptx 워크플로우입니다.

## Overview

새로운 PowerPoint 프레젠테이션을 처음부터 만들 때 사용합니다. HTML/CSS로 슬라이드를 디자인하고 PptxGenJS로 변환합니다.

## Workflow Decision Tree

```
새 프레젠테이션 생성?
├── 예 → 이 가이드 (html2pptx)
└── 아니오 (기존 템플릿 수정) → pptx-template.md
```

## Prerequisites

```bash
npm install -g pptxgenjs playwright sharp
```

---

## 워크플로우

```
1. 디자인 결정 (색상 팔레트, 폰트)
   ↓
2. HTML 슬라이드 작성 (각 슬라이드별 .html)
   ↓
3. PptxGenJS로 변환 (html2pptx)
   ↓
4. 차트/테이블 추가 (placeholder 영역)
   ↓
5. 썸네일 검증 및 수정
```

---

## Design Principles

**CRITICAL**: 프레젠테이션 생성 전 반드시 디자인 접근 방식을 결정하세요.

1. **주제 분석**: 프레젠테이션 주제가 무엇인가? 어떤 톤과 분위기가 적합한가?
2. **브랜딩 확인**: 회사/조직이 언급되면 브랜드 컬러 고려
3. **색상 팔레트 선택**: 주제에 맞는 3-5개 색상 선택
4. **접근 방식 설명**: 코드 작성 전 디자인 선택 이유 설명

### 웹 안전 폰트 (필수)

```
Arial, Helvetica, Times New Roman, Georgia, Courier New,
Verdana, Tahoma, Trebuchet MS, Impact
```

### 색상 팔레트 예시

| 이름 | 색상 코드 |
|------|-----------|
| Classic Blue | 1C2833, 2E4053, AAB7B8, F4F6F6 |
| Teal & Coral | 5EA8A7, 277884, FE4447, FFFFFF |
| Black & Gold | BF9A4A, 000000, F4F6F6 |
| Forest Green | 191A19, 4E9F3D, 1E5128, FFFFFF |
| Warm Blush | A49393, EED6D3, E8B4B8, FAF7F2 |
| Charcoal & Red | 292929, E33737, CCCBCB |
| Vibrant Orange | F96D00, F2F2F2, 222831 |
| Sage & Terracotta | 87A96B, E07A5F, F4F1DE, 2C2C2C |

---

## HTML 슬라이드 규칙

### 필수 치수

```css
/* 16:9 (기본) */
body {
  width: 720pt;
  height: 405pt;
  margin: 0;
  padding: 0;
}

/* 4:3 */
body { width: 720pt; height: 540pt; }

/* 16:10 */
body { width: 720pt; height: 450pt; }
```

### ⚠️ Critical Text Rules

**모든 텍스트는 반드시 `<p>`, `<h1>`-`<h6>`, `<ul>`, `<ol>` 안에!**

```html
<!-- ✅ CORRECT -->
<div><p>텍스트</p></div>
<h1>제목</h1>
<ul><li>항목</li></ul>

<!-- ❌ WRONG - 텍스트가 표시되지 않음 -->
<div>텍스트</div>
<span>텍스트</span>
```

### ⚠️ 절대 사용 금지

| 금지 항목 | 이유 | 대안 |
|----------|------|------|
| CSS gradient | 변환 안됨 | PNG 이미지로 대체 |
| 수동 bullet (•, -, *) | 잘못 표시됨 | `<ul>`, `<ol>` 사용 |
| 커스텀 폰트 | 렌더링 문제 | 웹 안전 폰트 사용 |

---

## ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `<div>텍스트</div>` | `<div><p>텍스트</p></div>` | 텍스트는 p 태그 안에 |
| `color: "#FF0000"` | `color: "FF0000"` | PptxGenJS에서 # 없이 |
| `linear-gradient(...)` | PNG 이미지 사용 | CSS gradient 불가 |
| `font-family: 'Roboto'` | `font-family: Arial` | 웹 안전 폰트만 |
| `• 항목 1` | `<ul><li>항목 1</li></ul>` | 수동 bullet 금지 |

---

## Complete HTML Example

```html
<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
  width: 720pt;
  height: 405pt;
  margin: 0;
  padding: 0;
  background: #1C2833;
  font-family: Arial, sans-serif;
  display: flex;
  justify-content: center;
  align-items: center;
}
.content {
  text-align: center;
  color: white;
}
h1 {
  font-size: 42pt;
  margin-bottom: 20pt;
  color: #FFFFFF;
}
p {
  font-size: 18pt;
  color: #AAB7B8;
}
</style>
</head>
<body>
<div class="content">
  <h1>프레젠테이션 제목</h1>
  <p>부제목 또는 설명</p>
</div>
</body>
</html>
```

---

## PptxGenJS 변환 코드

### 기본 변환

```javascript
const pptxgen = require('pptxgenjs');
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

async function html2pptx(htmlFile, pres) {
    const browser = await chromium.launch();
    const page = await browser.newPage();

    await page.setViewportSize({ width: 960, height: 540 });
    await page.goto(`file://${path.resolve(htmlFile)}`);

    // 스크린샷으로 변환
    const screenshot = await page.screenshot({ type: 'png' });
    await browser.close();

    const slide = pres.addSlide();
    slide.addImage({
        data: `data:image/png;base64,${screenshot.toString('base64')}`,
        x: 0, y: 0, w: '100%', h: '100%'
    });

    return { slide };
}

async function createPresentation() {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';
    pres.title = '프레젠테이션';
    pres.author = 'Author';

    // 슬라이드 추가
    await html2pptx('slide1.html', pres);
    await html2pptx('slide2.html', pres);
    await html2pptx('slide3.html', pres);

    await pres.writeFile('output.pptx');
    console.log('Presentation created: output.pptx');
}

createPresentation();
```

### Complete Multi-Slide Example

```javascript
const pptxgen = require('pptxgenjs');
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

async function html2pptx(htmlFile, pres) {
    const browser = await chromium.launch();
    const page = await browser.newPage();
    await page.setViewportSize({ width: 960, height: 540 });
    await page.goto(`file://${path.resolve(htmlFile)}`);
    const screenshot = await page.screenshot({ type: 'png' });
    await browser.close();

    const slide = pres.addSlide();
    slide.addImage({
        data: `data:image/png;base64,${screenshot.toString('base64')}`,
        x: 0, y: 0, w: '100%', h: '100%'
    });
    return { slide };
}

async function main() {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';
    pres.title = 'Project Presentation';

    // 슬라이드 파일 목록
    const slides = [
        'slides/01-title.html',
        'slides/02-overview.html',
        'slides/03-features.html',
        'slides/04-architecture.html',
        'slides/05-conclusion.html'
    ];

    for (const slideFile of slides) {
        if (fs.existsSync(slideFile)) {
            await html2pptx(slideFile, pres);
        }
    }

    await pres.writeFile('presentation.pptx');
}

main().catch(console.error);
```

---

## 차트 추가

### ⚠️ CRITICAL: 색상에 # 사용 금지!

```javascript
// ❌ WRONG - 파일 손상됨
chartColors: ["#4472C4", "#ED7D31"]

// ✅ CORRECT
chartColors: ["4472C4", "ED7D31"]
```

### Bar Chart

```javascript
slide.addChart(pres.charts.BAR, [{
    name: "Sales 2024",
    labels: ["Q1", "Q2", "Q3", "Q4"],
    values: [4500, 5500, 6200, 7100]
}], {
    x: 1, y: 1.5, w: 8, h: 4,
    barDir: 'col',
    showTitle: true,
    title: 'Quarterly Sales',
    showLegend: false,
    showCatAxisTitle: true,
    catAxisTitle: 'Quarter',
    showValAxisTitle: true,
    valAxisTitle: 'Sales ($000s)',
    chartColors: ["4472C4"]
});
```

### Line Chart

```javascript
slide.addChart(pres.charts.LINE, [{
    name: "Revenue",
    labels: ["Jan", "Feb", "Mar", "Apr", "May"],
    values: [100, 120, 115, 140, 160]
}], {
    x: 1, y: 1.5, w: 8, h: 4,
    lineSize: 3,
    lineSmooth: true,
    showCatAxisTitle: true,
    catAxisTitle: 'Month',
    showValAxisTitle: true,
    valAxisTitle: 'Revenue ($M)',
    chartColors: ["4472C4", "ED7D31"]
});
```

### Pie Chart

```javascript
slide.addChart(pres.charts.PIE, [{
    name: "Market Share",
    labels: ["Product A", "Product B", "Product C"],
    values: [45, 35, 20]
}], {
    x: 2, y: 1, w: 6, h: 4,
    showPercent: true,
    showLegend: true,
    legendPos: 'r',
    chartColors: ["4472C4", "ED7D31", "A5A5A5"]
});
```

---

## 테이블 추가

### Basic Table

```javascript
slide.addTable([
    ["Header 1", "Header 2", "Header 3"],
    ["Row 1", "Value 1", "100"],
    ["Row 2", "Value 2", "200"]
], {
    x: 1, y: 1.5, w: 8, h: 3,
    border: { pt: 1, color: "CCCCCC" },
    fontSize: 14
});
```

### Styled Table

```javascript
const tableData = [
    [
        { text: "항목", options: { fill: { color: "4472C4" }, color: "FFFFFF", bold: true } },
        { text: "값", options: { fill: { color: "4472C4" }, color: "FFFFFF", bold: true } }
    ],
    ["제품 A", "$50,000"],
    ["제품 B", "$35,000"],
    ["제품 C", "$28,000"]
];

slide.addTable(tableData, {
    x: 1, y: 2, w: 8, h: 3,
    colW: [4, 4],
    border: { pt: 1, color: "CCCCCC" },
    align: "center",
    valign: "middle"
});
```

---

## 그라데이션/아이콘 처리

CSS 그라데이션은 PowerPoint로 변환되지 않습니다. PNG 이미지로 먼저 생성하세요.

### 그라데이션 배경 생성

```javascript
const sharp = require('sharp');

async function createGradientBackground(filename, color1, color2) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="563">
        <defs>
            <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" style="stop-color:#${color1}"/>
                <stop offset="100%" style="stop-color:#${color2}"/>
            </linearGradient>
        </defs>
        <rect width="100%" height="100%" fill="url(#g)"/>
    </svg>`;

    await sharp(Buffer.from(svg)).png().toFile(filename);
    return filename;
}

// 사용
await createGradientBackground('bg.png', '1C2833', '2E4053');
```

### 아이콘 생성

```javascript
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const { FaHome } = require('react-icons/fa');

async function rasterizeIcon(IconComponent, color, size, filename) {
    const svgString = ReactDOMServer.renderToStaticMarkup(
        React.createElement(IconComponent, { color: `#${color}`, size: size })
    );
    await sharp(Buffer.from(svgString)).png().toFile(filename);
    return filename;
}

// 사용
await rasterizeIcon(FaHome, "4472C4", "256", "home-icon.png");
```

---

## 레이아웃 가이드

### 차트/테이블 포함 슬라이드

**권장**: 2단 레이아웃 (텍스트 + 차트)

```
┌────────────────────────────────┐
│          Header Title          │
├───────────────┬────────────────┤
│               │                │
│  Bullet List  │     Chart      │
│               │                │
│               │                │
└───────────────┴────────────────┘
```

**피하기**: 세로 스택 (텍스트 아래 차트)

```
# ❌ Bad Layout
┌────────────────────────────────┐
│          Header Title          │
├────────────────────────────────┤
│  Bullet List                   │
├────────────────────────────────┤
│     Chart (too small)          │
└────────────────────────────────┘
```

---

## 썸네일 검증

### 썸네일 생성

```bash
# PDF 변환 후 이미지로
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

### 검증 체크리스트

- [ ] 텍스트가 잘리지 않았는지
- [ ] 텍스트가 겹치지 않는지
- [ ] 콘텐츠가 슬라이드 경계에 너무 가깝지 않은지
- [ ] 배경과 텍스트 대비가 충분한지
- [ ] 차트/테이블이 읽을 수 있는 크기인지

---

## Troubleshooting

### "Text not appearing"
- **원인**: 텍스트가 `<p>`, `<h1>`-`<h6>`, `<ul>`, `<ol>` 밖에 있음
- **해결**: 모든 텍스트를 적절한 태그로 감싸기

### "File corrupted"
- **원인**: 색상 코드에 `#` 포함
- **해결**: PptxGenJS에서 `#` 제거 (`"#FF0000"` → `"FF0000"`)

### "Gradient not showing"
- **원인**: CSS gradient는 변환되지 않음
- **해결**: Sharp로 PNG 이미지 생성 후 사용

### "Font rendering issue"
- **원인**: 커스텀 폰트 사용
- **해결**: 웹 안전 폰트만 사용 (Arial, Helvetica 등)

### "Chart colors wrong"
- **원인**: chartColors 배열에 `#` 포함
- **해결**: `["4472C4"]` 형식 사용

---

## Quick Reference

| 작업 | 코드 |
|------|------|
| 슬라이드 추가 | `pres.addSlide()` |
| 이미지 추가 | `slide.addImage({path: "img.png", x, y, w, h})` |
| 텍스트 추가 | `slide.addText("text", {x, y, w, h})` |
| 차트 추가 | `slide.addChart(pres.charts.BAR, data, opts)` |
| 테이블 추가 | `slide.addTable(rows, opts)` |
| 저장 | `pres.writeFile("output.pptx")` |

### 레이아웃 상수

```javascript
pres.layout = 'LAYOUT_16x9';  // 16:9
pres.layout = 'LAYOUT_4x3';   // 4:3
pres.layout = 'LAYOUT_16x10'; // 16:10
```

### 차트 유형

```javascript
pres.charts.BAR      // 막대 차트
pres.charts.LINE     // 선 차트
pres.charts.PIE      // 파이 차트
pres.charts.DOUGHNUT // 도넛 차트
pres.charts.SCATTER  // 산점도
```
