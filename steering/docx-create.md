# Word 문서 생성 (JavaScript docx)

docx 라이브러리를 사용한 새 Word 문서 생성 가이드입니다.

## Overview

JavaScript/TypeScript로 .docx 파일을 프로그래밍 방식으로 생성합니다. 레이아웃, 스타일, 테이블, 이미지 등을 코드로 제어할 수 있습니다. 생성 후 반드시 검증합니다.

## Workflow Decision Tree

```
Word 문서 작업 유형?
├── 새 문서 생성 → 이 가이드 (docx.js)
├── 기존 문서 편집 → docx-edit.md (OOXML)
└── 기존 문서 텍스트 추출 → pandoc 사용
```

## Prerequisites

```bash
npm install -g docx
```

---

## ⚠️ CRITICAL: 자주 발생하는 오류

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `new TextRun("Line 1\nLine 2")` | 별도 Paragraph로 분리 | `\n` 줄바꿈 금지 |
| `new Paragraph({text: "..."})` | `children: [new TextRun(...)]` | TextRun 필수 |
| `new PageBreak()` | `new Paragraph({children: [new PageBreak()]})` | PageBreak는 Paragraph 안에 |
| `"• Item"` 또는 `SymbolRun` | `numbering` config 사용 | 유니코드 bullet 금지 |
| `ShadingType.SOLID` | `ShadingType.CLEAR` | 테이블 셀 배경 |
| `color: "#FF0000"` | `color: "FF0000"` | # 없이 |
| ImageRun type 미지정 | `type: "png"` 필수 | 이미지 타입 명시 |
| `WidthType.PERCENTAGE` | `WidthType.DXA` | Google Docs 호환 |
| 테이블을 divider로 사용 | Paragraph border 사용 | 셀 최소 높이 문제 |
| 페이지 크기 미지정 (A4 기본) | US Letter 명시 설정 | 12240 x 15840 DXA |

---

## 기본 문서 구조

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        InternalHyperlink, Bookmark, FootnoteReferenceRun, PositionalTab,
        PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
        TabStopType, TabStopPosition, Column, SectionType,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');
const fs = require('fs');

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },  // US Letter (DXA)
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [/* 콘텐츠 */]
  }]
});

Packer.toBuffer(doc).then(buffer => fs.writeFileSync("output.docx", buffer));
```

### 검증 (필수)

```bash
soffice --headless --convert-to pdf doc.docx  # 열리면 유효
```

검증 실패 시 unpack → XML 수정 → repack.

### 페이지 크기 (DXA 단위, 1440 = 1인치)

| 용지 | Width | Height | Content Width (1" 마진) |
|------|-------|--------|------------------------|
| US Letter | 12,240 | 15,840 | 9,360 |
| A4 (기본값) | 11,906 | 16,838 | 9,026 |

### 가로 방향 (Landscape)

docx-js가 width/height를 내부적으로 교환하므로, 세로 치수를 그대로 전달:

```javascript
size: {
  width: 12240,   // 짧은 변을 width로
  height: 15840,  // 긴 변을 height로
  orientation: PageOrientation.LANDSCAPE  // docx-js가 XML에서 교환
},
```

---

## 스타일 정의

Arial을 기본 폰트로 사용 (범용 호환). 제목은 검정색으로 가독성 유지.

```javascript
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },  // 12pt 기본
    paragraphStyles: [
      // IMPORTANT: 내장 스타일 오버라이드시 정확한 ID 사용
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },  // TOC 필수
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Title")] }),
    ]
  }]
});
```

---

## 텍스트 & 포맷팅

```javascript
// ⚠️ CRITICAL: 줄바꿈은 별도 Paragraph로!
new Paragraph({ children: [new TextRun("Line 1")] }),
new Paragraph({ children: [new TextRun("Line 2")] })

// 텍스트 스타일
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 200 },
  children: [
    new TextRun({ text: "Bold", bold: true }),
    new TextRun({ text: "Italic", italics: true }),
    new TextRun({ text: "Colored", color: "FF0000", size: 28, font: "Arial" }),
    new TextRun({ text: "Highlighted", highlight: "yellow" }),
    new TextRun({ text: "x2", superScript: true }),
  ]
})
```

---

## 리스트 (⚠️ 유니코드 bullet 금지)

```javascript
const doc = new Document({
  numbering: {
    config: [
      // Bullet 리스트
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      // 번호 리스트
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Numbered item")] }),
    ]
  }]
});

// ⚠️ 번호 규칙:
// 같은 reference = 번호 연속 (1,2,3 → 4,5,6)
// 다른 reference = 번호 재시작 (1,2,3 → 1,2,3)
```

---

## 테이블

**CRITICAL: 이중 너비 설정 필수** — `columnWidths` + 각 셀 `width` 모두 설정.

```javascript
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

new Table({
  width: { size: 9360, type: WidthType.DXA },  // 항상 DXA (% 금지 — Google Docs 깨짐)
  columnWidths: [4680, 4680],  // 합계 = 테이블 width
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          borders,
          width: { size: 4680, type: WidthType.DXA },
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR },  // ⚠️ CLEAR 필수!
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Header", bold: true })]
          })]
        }),
        new TableCell({
          borders,
          width: { size: 4680, type: WidthType.DXA },
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun("Header 2")] })]
        })
      ]
    })
  ]
})
```

**너비 규칙:**
- 항상 `WidthType.DXA` 사용 — `WidthType.PERCENTAGE` 금지 (Google Docs 비호환)
- 테이블 width = `columnWidths` 합계
- 셀 `width` = 해당 `columnWidth`와 일치
- 셀 `margins`는 내부 패딩 — 콘텐츠 영역 축소, 셀 너비에 추가 아님
- 전체 너비 테이블: 콘텐츠 너비 = 페이지 너비 - 좌우 마진

**⚠️ 테이블을 구분선으로 사용 금지** — 셀에 최소 높이가 있어 빈 박스로 렌더링됨. 대신 Paragraph border 사용:
```javascript
new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } }
})
```

---

## 이미지

```javascript
// ⚠️ CRITICAL: type 파라미터 필수!
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new ImageRun({
    type: "png",  // 필수: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" }  // 세 필드 모두 필수
  })]
})
```

---

## 하이퍼링크

```javascript
// 외부 링크
new Paragraph({
  children: [new ExternalHyperlink({
    children: [new TextRun({ text: "Click here", style: "Hyperlink" })],
    link: "https://example.com",
  })]
})

// 내부 링크 (북마크 + 참조)
// 1. 대상에 북마크 생성
new Paragraph({ heading: HeadingLevel.HEADING_1, children: [
  new Bookmark({ id: "chapter1", children: [new TextRun("Chapter 1")] }),
]})
// 2. 링크 생성
new Paragraph({ children: [new InternalHyperlink({
  children: [new TextRun({ text: "See Chapter 1", style: "Hyperlink" })],
  anchor: "chapter1",
})]})
```

---

## 각주

```javascript
const doc = new Document({
  footnotes: {
    1: { children: [new Paragraph("Source: Annual Report 2024")] },
    2: { children: [new Paragraph("See appendix for methodology")] },
  },
  sections: [{
    children: [new Paragraph({
      children: [
        new TextRun("Revenue grew 15%"),
        new FootnoteReferenceRun(1),
        new TextRun(" using adjusted metrics"),
        new FootnoteReferenceRun(2),
      ],
    })]
  }]
});
```

---

## 탭 정지

```javascript
// 같은 줄에 우측 정렬 (예: 제목 반대편에 날짜)
new Paragraph({
  children: [
    new TextRun("Company Name"),
    new TextRun("\tJanuary 2025"),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
})

// 점선 리더 (목차 스타일)
new Paragraph({
  children: [
    new TextRun("Introduction"),
    new TextRun({ children: [
      new PositionalTab({
        alignment: PositionalTabAlignment.RIGHT,
        relativeTo: PositionalTabRelativeTo.MARGIN,
        leader: PositionalTabLeader.DOT,
      }),
      "3",
    ]}),
  ],
})
```

---

## 다단 레이아웃

```javascript
// 동일 너비 컬럼
sections: [{
  properties: {
    column: {
      count: 2,
      space: 720,        // 컬럼 간격 (DXA, 720 = 0.5인치)
      equalWidth: true,
      separate: true,    // 컬럼 사이 세로선
    },
  },
  children: [/* 콘텐츠가 자연스럽게 컬럼을 넘김 */]
}]

// 커스텀 너비 (equalWidth: false 필수)
sections: [{
  properties: {
    column: {
      equalWidth: false,
      children: [
        new Column({ width: 5400, space: 720 }),
        new Column({ width: 3240 }),
      ],
    },
  },
  children: [/* 콘텐츠 */]
}]
```

컬럼 강제 분리: `type: SectionType.NEXT_COLUMN` 섹션 사용.

---

## 페이지 나누기

```javascript
// ⚠️ CRITICAL: PageBreak는 반드시 Paragraph 안에!
new Paragraph({ children: [new PageBreak()] })

// 또는 pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

---

## 목차

```javascript
// ⚠️ CRITICAL: HeadingLevel만 사용, custom style 금지
new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })
```

---

## 헤더 & 푸터

```javascript
sections: [{
  properties: {
    page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

**⚠️ 푸터에 테이블 사용 금지** — 빈 박스 렌더링. 2단 푸터는 탭 정지 사용.

---

## Quick Reference

### 측정 단위

- DXA (1/20 포인트): 1440 = 1인치
- Size (반포인트): 24 = 12pt

### 테이블 열 너비 (US Letter, 1" 마진 = 9360 DXA)

- 2열: [4680, 4680]
- 3열: [3120, 3120, 3120]

### 핵심 규칙 요약

- 페이지 크기 명시 설정 (기본 A4)
- Landscape: 세로 치수 전달, docx-js가 교환
- `\n` 금지 — Paragraph 분리
- 유니코드 bullet 금지 — `LevelFormat.BULLET`
- PageBreak는 Paragraph 안에
- ImageRun에 `type` 필수
- 테이블은 `WidthType.DXA`만 사용
- 테이블 이중 너비 (`columnWidths` + 셀 `width`)
- `ShadingType.CLEAR` 사용 (SOLID 금지)
- 테이블을 구분선으로 사용 금지
- TOC는 `HeadingLevel`만
- 내장 스타일 오버라이드: 정확한 ID ("Heading1" 등)
- `outlineLevel` 필수 (TOC용)

---

## Troubleshooting

### "Invalid XML" 또는 Word에서 열리지 않음
- **원인**: PageBreak가 Paragraph 밖에 있음
- **해결**: `new Paragraph({ children: [new PageBreak()] })`

### 테이블 셀이 검은색 배경
- **원인**: `ShadingType.SOLID` 사용
- **해결**: `ShadingType.CLEAR` 사용

### 테이블 너비가 Google Docs에서 깨짐
- **원인**: `WidthType.PERCENTAGE` 사용
- **해결**: `WidthType.DXA` 사용

### TOC가 동작하지 않음
- **원인**: heading에 custom style 추가 또는 outlineLevel 누락
- **해결**: `heading: HeadingLevel.HEADING_1`만 사용, 스타일에 `outlineLevel` 포함

### 이미지 삽입 오류
- **원인**: `type` 파라미터 누락
- **해결**: `type: "png"` 추가

### 리스트 bullet이 이상하게 표시됨
- **원인**: 유니코드 bullet 또는 SymbolRun 사용
- **해결**: numbering config 사용
