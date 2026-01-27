# Word 문서 생성 (JavaScript docx)

docx 라이브러리를 사용한 새 Word 문서 생성 가이드입니다.

## Overview

JavaScript/TypeScript로 .docx 파일을 프로그래밍 방식으로 생성합니다. 레이아웃, 스타일, 테이블, 이미지 등을 코드로 제어할 수 있습니다.

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

### ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `new TextRun("Line 1\nLine 2")` | 별도 Paragraph로 분리 | `\n` 줄바꿈 금지 |
| `new Paragraph({text: "..."})` | `children: [new TextRun(...)]` | TextRun 필수 |
| `new PageBreak()` | `new Paragraph({children: [new PageBreak()]})` | PageBreak는 Paragraph 안에 |
| `format: "bullet"` | `format: LevelFormat.BULLET` | 상수 사용 필수 |
| `"• Item"` 또는 `SymbolRun` | `numbering` config 사용 | 유니코드 bullet 금지 |
| `ShadingType.SOLID` | `ShadingType.CLEAR` | 테이블 셀 배경 |
| `color: "#FF0000"` | `color: "FF0000"` | # 없이 |
| ImageRun type 미지정 | `type: "png"` 필수 | 이미지 타입 명시 |

---

## 기본 문서 구조

```javascript
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        Table, TableRow, TableCell, ImageRun, Header, Footer, PageNumber,
        LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
        ExternalHyperlink, TableOfContents, PageBreak, UnderlineType } = require('docx');
const fs = require('fs');

// 기본 문서 생성
const doc = new Document({
    sections: [{
        properties: {
            page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
        },
        children: [/* 콘텐츠 */]
    }]
});

// 저장
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("output.docx", buffer));
```

---

## 텍스트 & 포맷팅

### 기본 텍스트

```javascript
// ⚠️ CRITICAL: 줄바꿈은 별도 Paragraph로!
// ❌ WRONG: new TextRun("Line 1\nLine 2")
// ✅ CORRECT:
new Paragraph({ children: [new TextRun("Line 1")] }),
new Paragraph({ children: [new TextRun("Line 2")] })
```

### 텍스트 스타일

```javascript
new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
    indent: { left: 720, right: 720 },
    children: [
        new TextRun({ text: "Bold", bold: true }),
        new TextRun({ text: "Italic", italics: true }),
        new TextRun({ text: "Underlined", underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: "Colored", color: "FF0000", size: 28, font: "Arial" }),
        new TextRun({ text: "Highlighted", highlight: "yellow" }),
        new TextRun({ text: "Strikethrough", strike: true }),
        new TextRun({ text: "x2", superScript: true }),
        new TextRun({ text: "H2O", subScript: true })
    ]
})
```

---

## 스타일 정의

```javascript
const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 24 } } },  // 12pt 기본
        paragraphStyles: [
            // 제목 스타일 (기본 Title 오버라이드)
            { id: "Title", name: "Title", basedOn: "Normal",
              run: { size: 56, bold: true, color: "000000", font: "Arial" },
              paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER } },
            // Heading1 (TOC에 필요)
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
              run: { size: 32, bold: true, color: "000000", font: "Arial" },
              paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },  // TOC 필수
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
              run: { size: 28, bold: true, color: "000000", font: "Arial" },
              paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } }
        ]
    },
    sections: [{
        children: [
            new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Document Title")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Section 1")] })
        ]
    }]
});
```

**전문적인 폰트 조합:**
- Arial (Headers) + Arial (Body) - 가장 범용적
- Times New Roman (Headers) + Arial (Body) - 전통적
- Georgia (Headers) + Verdana (Body) - 화면 최적화

---

## 리스트 (⚠️ CRITICAL: 유니코드 bullet 금지)

```javascript
// ⚠️ CRITICAL: 유니코드 bullet 사용 금지!
// ❌ WRONG: new TextRun("• Item")
// ❌ WRONG: new SymbolRun({ char: "2022" })
// ✅ CORRECT: numbering config 사용

const doc = new Document({
    numbering: {
        config: [
            // Bullet 리스트
            { reference: "bullet-list",
              levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            // 번호 리스트 1
            { reference: "first-numbered-list",
              levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            // 번호 리스트 2 (다른 reference = 1부터 다시 시작)
            { reference: "second-numbered-list",
              levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
        ]
    },
    sections: [{
        children: [
            // Bullet 리스트
            new Paragraph({ numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("First bullet")] }),
            new Paragraph({ numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Second bullet")] }),
            // 번호 리스트
            new Paragraph({ numbering: { reference: "first-numbered-list", level: 0 },
                children: [new TextRun("First item")] }),  // 1.
            new Paragraph({ numbering: { reference: "first-numbered-list", level: 0 },
                children: [new TextRun("Second item")] })  // 2.
        ]
    }]
});

// ⚠️ CRITICAL 번호 규칙:
// - 같은 reference = 번호 연속 (1, 2, 3... 이후 4, 5, 6...)
// - 다른 reference = 번호 재시작 (1, 2, 3... 이후 1, 2, 3...)
```

---

## 테이블

```javascript
// ⚠️ CRITICAL: columnWidths + 각 셀 width 둘 다 설정
// ⚠️ CRITICAL: ShadingType.CLEAR 사용 (SOLID 금지)

const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };

new Table({
    columnWidths: [4680, 4680],  // 값: DXA (1440 = 1인치)
    margins: { top: 100, bottom: 100, left: 180, right: 180 },
    rows: [
        new TableRow({
            tableHeader: true,
            children: [
                new TableCell({
                    borders: cellBorders,
                    width: { size: 4680, type: WidthType.DXA },
                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },  // ⚠️ CLEAR 필수!
                    verticalAlign: VerticalAlign.CENTER,
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "Header", bold: true })]
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 4680, type: WidthType.DXA },
                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                    children: [new Paragraph({ children: [new TextRun("Header 2")] })]
                })
            ]
        }),
        new TableRow({
            children: [
                new TableCell({
                    borders: cellBorders,
                    width: { size: 4680, type: WidthType.DXA },
                    children: [new Paragraph({ children: [new TextRun("Data 1")] })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 4680, type: WidthType.DXA },
                    children: [
                        // 셀 내 bullet 리스트
                        new Paragraph({ numbering: { reference: "bullet-list", level: 0 },
                            children: [new TextRun("Bullet in cell")] })
                    ]
                })
            ]
        })
    ]
})

// 열 너비 참조 (Letter size, 1인치 마진 = 9360 DXA):
// - 2열: [4680, 4680]
// - 3열: [3120, 3120, 3120]
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
        transformation: { width: 200, height: 150, rotation: 0 },
        altText: { title: "Logo", description: "Company logo", name: "Logo" }  // 세 필드 모두 필수
    })]
})
```

---

## 링크 & 목차

```javascript
// Table of Contents (⚠️ HeadingLevel만 사용, custom style 금지)
new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),

// ❌ WRONG: heading + style 같이 사용
// new Paragraph({ heading: HeadingLevel.HEADING_1, style: "customHeader", children: [...] })
// ✅ CORRECT:
new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Section")] })

// 외부 링크
new Paragraph({
    children: [new ExternalHyperlink({
        children: [new TextRun({ text: "Google", style: "Hyperlink" })],
        link: "https://www.google.com"
    })]
})
```

---

## 페이지 나누기

```javascript
// ⚠️ CRITICAL: PageBreak는 반드시 Paragraph 안에!
// ❌ WRONG: new PageBreak()
// ✅ CORRECT:
new Paragraph({ children: [new PageBreak()] })

// 또는 pageBreakBefore 사용
new Paragraph({
    pageBreakBefore: true,
    children: [new TextRun("This starts on a new page")]
})
```

---

## 헤더 & 푸터

```javascript
const doc = new Document({
    sections: [{
        properties: {
            page: {
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                pageNumbers: { start: 1, formatType: "decimal" }
            }
        },
        headers: {
            default: new Header({ children: [new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun("Header Text")]
            })] })
        },
        footers: {
            default: new Footer({ children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun("Page "),
                    new TextRun({ children: [PageNumber.CURRENT] }),
                    new TextRun(" of "),
                    new TextRun({ children: [PageNumber.TOTAL_PAGES] })
                ]
            })] })
        },
        children: [/* content */]
    }]
});
```

---

## 완전한 예제

```javascript
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        Table, TableRow, TableCell, ImageRun, Header, Footer, PageNumber,
        LevelFormat, BorderStyle, WidthType, ShadingType, PageBreak } = require('docx');
const fs = require('fs');

async function createProfessionalDocument() {
    const doc = new Document({
        styles: {
            default: { document: { run: { font: "Arial", size: 24 } } },
            paragraphStyles: [
                { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
                  run: { size: 32, bold: true }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } }
            ]
        },
        numbering: {
            config: [
                { reference: "bullets", levels: [
                    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }
                ] }
            ]
        },
        sections: [{
            properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
            headers: {
                default: new Header({ children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Company Name", bold: true })]
                })] })
            },
            footers: {
                default: new Footer({ children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
                })] })
            },
            children: [
                new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Project Report")] }),
                new Paragraph({ children: [new TextRun("This document describes the project details.")] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 },
                    children: [new TextRun("First item")] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 },
                    children: [new TextRun("Second item")] }),
                new Paragraph({ children: [new PageBreak()] }),
                new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Section 2")] })
            ]
        }]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("professional_report.docx", buffer);
}

createProfessionalDocument();
```

---

## Quick Reference

### 측정 단위

- DXA (twentieths of a point): 1440 = 1인치
- Size (half-points): 24 = 12pt

### 밑줄 스타일

`SINGLE`, `DOUBLE`, `WAVY`, `DASH`

### 테두리 스타일

`SINGLE`, `DOUBLE`, `DASHED`, `DOTTED`

### 번호 형식

`DECIMAL` (1,2,3), `UPPER_ROMAN` (I,II,III), `LOWER_LETTER` (a,b,c)

### 하이라이트 색상

`yellow`, `green`, `cyan`, `magenta`, `blue`, `red`, `darkBlue`, `darkCyan`, `darkGreen`, `darkMagenta`, `darkRed`, `darkYellow`, `darkGray`, `lightGray`, `black`

---

## Troubleshooting

### "Invalid XML" 또는 Word에서 열리지 않음
- **원인**: PageBreak가 Paragraph 밖에 있음
- **해결**: `new Paragraph({ children: [new PageBreak()] })`

### 테이블 셀이 검은색 배경
- **원인**: `ShadingType.SOLID` 사용
- **해결**: `ShadingType.CLEAR` 사용

### TOC가 동작하지 않음
- **원인**: heading에 custom style 추가
- **해결**: `heading: HeadingLevel.HEADING_1`만 사용, style 제거

### 이미지 삽입 오류
- **원인**: `type` 파라미터 누락
- **해결**: `type: "png"` 추가

### 리스트 bullet이 이상하게 표시됨
- **원인**: 유니코드 bullet(•) 또는 SymbolRun 사용
- **해결**: numbering config 사용
