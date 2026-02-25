# PDF 텍스트/테이블 추출 및 조작

PDF 처리의 핵심 가이드입니다. 폼 작성은 `pdf-forms.md`를 참조하세요.

## Overview

PDF 파일에서 텍스트/테이블 추출, 병합/분할, 워터마크, 이미지 변환, 암호화를 다룹니다.

## Workflow Decision Tree

```
PDF 작업 유형?
├── 텍스트 추출 → pdfplumber (권장) 또는 pypdf
├── 테이블 추출 → pdfplumber + pandas
├── PDF 병합/분할 → pypdf 또는 qpdf
├── 워터마크 추가 → pypdf + reportlab
├── PDF → 이미지 → pdftoppm 또는 pypdfium2
├── 이미지 추출 → pdfimages (poppler)
├── PDF 생성 → reportlab
├── 폼 채우기 → pdf-forms.md
├── OCR (스캔 PDF) → pytesseract + pdf2image
├── 암호화/복호화 → pypdf 또는 qpdf
└── 메타데이터 수정 → pypdf
```

## Prerequisites

```bash
pip install pypdf pdfplumber pandas reportlab

# 이미지 변환용
brew install poppler  # macOS (pdftoppm, pdfimages)

# 고급 조작
brew install qpdf

# OCR용
pip install pytesseract pdf2image
```

---

## Quick Start

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("document.pdf")
print(f"Pages: {len(reader.pages)}")

text = ""
for page in reader.pages:
    text += page.extract_text()
```

---

## 텍스트 추출

### pdfplumber (권장 — 레이아웃 기반)

```python
import pdfplumber

with pdfplumber.open("document.pdf") as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            print(text)
```

### 특정 영역 추출

```python
with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]
    # bbox: (x0, y0, x1, y1) — 왼쪽 상단 기준
    cropped = page.within_bbox((0, 0, page.width/2, page.height/2))
    text = cropped.extract_text()
```

### pypdf (더 빠름, 레이아웃 덜 정확)

```python
from pypdf import PdfReader

reader = PdfReader("document.pdf")
for page in reader.pages:
    print(page.extract_text())
```

### 커맨드라인

```bash
# 텍스트 추출
pdftotext input.pdf output.txt

# 레이아웃 보존
pdftotext -layout input.pdf output.txt

# 특정 페이지 (1-5)
pdftotext -f 1 -l 5 input.pdf output.txt
```

---

## 테이블 추출

### 기본

```python
import pdfplumber
import pandas as pd

with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]
    tables = page.extract_tables()

    for i, table in enumerate(tables):
        if table and len(table) > 1:
            df = pd.DataFrame(table[1:], columns=table[0])
            df.to_csv(f"table_{i+1}.csv", index=False)
```

### 커스텀 설정 (복잡한 레이아웃)

```python
# 선 기반 (기본)
table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
    "intersection_tolerance": 15
}

# 텍스트 기반 (선 없는 테이블)
table_settings = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
}

tables = page.extract_tables(table_settings)
```

### 모든 페이지에서 추출

```python
def extract_all_tables(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            for table in page.extract_tables():
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df['source_page'] = page_num + 1
                    all_tables.append(df)
    if all_tables:
        return pd.concat(all_tables, ignore_index=True)
    return None
```

---

## PDF 병합

```python
from pypdf import PdfWriter, PdfReader

writer = PdfWriter()
for pdf_file in ["doc1.pdf", "doc2.pdf", "doc3.pdf"]:
    reader = PdfReader(pdf_file)
    for page in reader.pages:
        writer.add_page(page)

with open("merged.pdf", "wb") as output:
    writer.write(output)
```

### 커맨드라인

```bash
# qpdf
qpdf --empty --pages doc1.pdf doc2.pdf -- merged.pdf

# 특정 페이지만
qpdf --empty --pages doc1.pdf 1-3 doc2.pdf 5-7 -- combined.pdf
```

---

## PDF 분할

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("input.pdf")
for i, page in enumerate(reader.pages):
    writer = PdfWriter()
    writer.add_page(page)
    with open(f"page_{i+1}.pdf", "wb") as output:
        writer.write(output)
```

### 페이지 범위 추출

```python
def extract_pages(input_path, output_path, start, end):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for i in range(start - 1, min(end, len(reader.pages))):
        writer.add_page(reader.pages[i])
    with open(output_path, "wb") as f:
        writer.write(f)

extract_pages("document.pdf", "pages_3_to_5.pdf", 3, 5)
```

### 커맨드라인

```bash
qpdf input.pdf --pages . 1-5 -- pages1-5.pdf
qpdf --split-pages=3 input.pdf output_%02d.pdf  # 3페이지씩 분할
```

---

## PDF 생성 (reportlab)

### 기본

```python
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

c = canvas.Canvas("hello.pdf", pagesize=letter)
width, height = letter
c.drawString(100, height - 100, "Hello World!")
c.line(100, height - 140, 400, height - 140)
c.save()
```

### 전문 보고서 (Platypus)

```python
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

doc = SimpleDocTemplate("report.pdf", pagesize=letter)
styles = getSampleStyleSheet()
story = []

story.append(Paragraph("Report Title", styles['Title']))
story.append(Spacer(1, 12))
story.append(Paragraph("Body text content. " * 20, styles['Normal']))
story.append(PageBreak())
story.append(Paragraph("Page 2", styles['Heading1']))

doc.build(story)
```

### ⚠️ 아래첨자/위첨자

**유니코드 아래첨자/위첨자 문자 사용 금지** (내장 폰트에 글리프 없음 — 검은 박스로 렌더링).

```python
# ✅ CORRECT: XML 마크업 태그 사용
chemical = Paragraph("H<sub>2</sub>O", styles['Normal'])
squared = Paragraph("x<super>2</super>", styles['Normal'])

# ❌ WRONG: 유니코드 문자 사용
# Paragraph("H₂O", ...)  # 검은 박스
```

---

## 워터마크

```python
from pypdf import PdfReader, PdfWriter

watermark = PdfReader("watermark.pdf").pages[0]
reader = PdfReader("document.pdf")
writer = PdfWriter()

for page in reader.pages:
    page.merge_page(watermark)
    writer.add_page(page)

with open("watermarked.pdf", "wb") as output:
    writer.write(output)
```

---

## PDF → 이미지 변환

### pdftoppm (권장)

```bash
pdftoppm -jpeg -r 150 document.pdf output      # 전체
pdftoppm -png -r 300 -f 1 -l 3 document.pdf pages  # 특정 페이지
```

### pypdfium2

```python
import pypdfium2 as pdfium

pdf = pdfium.PdfDocument("document.pdf")
for i, page in enumerate(pdf):
    bitmap = page.render(scale=2.0)
    img = bitmap.to_pil()
    img.save(f"page_{i+1}.png", "PNG")
```

---

## 이미지 추출

```bash
# 내장 이미지 추출 (poppler)
pdfimages -j input.pdf output_prefix
pdfimages -all input.pdf images/img

# 이미지 목록 확인
pdfimages -list document.pdf
```

---

## 페이지 회전

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("input.pdf")
writer = PdfWriter()

page = reader.pages[0]
page.rotate(90)  # 시계 방향 90도
writer.add_page(page)

with open("rotated.pdf", "wb") as output:
    writer.write(output)
```

---

## 암호화/복호화

```python
from pypdf import PdfReader, PdfWriter

# 암호화
reader = PdfReader("input.pdf")
writer = PdfWriter()
for page in reader.pages:
    writer.add_page(page)
writer.encrypt("userpassword", "ownerpassword")
with open("encrypted.pdf", "wb") as output:
    writer.write(output)
```

### 커맨드라인

```bash
# 암호화
qpdf --encrypt user_pass owner_pass 256 --print=none --modify=none -- input.pdf encrypted.pdf

# 복호화
qpdf --password=secret --decrypt encrypted.pdf decrypted.pdf

# 암호화 상태 확인
qpdf --show-encryption encrypted.pdf
```

---

## OCR (스캔 PDF)

```python
import pytesseract
from pdf2image import convert_from_path

images = convert_from_path('scanned.pdf')
for i, image in enumerate(images):
    text = pytesseract.image_to_string(image, lang='eng')
    print(f"Page {i+1}: {text}")
```

---

## 메타데이터

```python
from pypdf import PdfReader, PdfWriter

# 조회
reader = PdfReader("document.pdf")
if reader.metadata:
    for key, value in reader.metadata.items():
        print(f"{key}: {value}")

# 수정
writer = PdfWriter()
writer.append(reader)
writer.add_metadata({
    "/Title": "Updated Title",
    "/Author": "Author Name",
})
with open("updated.pdf", "wb") as f:
    writer.write(f)
```

---

## PDF 수리/최적화 (qpdf)

```bash
# 구조 검사
qpdf --check input.pdf

# 수리
qpdf --replace-input corrupted.pdf

# 웹 최적화 (스트리밍)
qpdf --linearize input.pdf optimized.pdf

# 압축
qpdf --optimize-level=all input.pdf compressed.pdf
```

---

## 좌표 체계

```
pdfplumber (왼쪽 상단 기준)         reportlab (왼쪽 하단 기준)
(0, 0) ┌──────────────┐            ┌──────────────┐ (0, height)
       │              │            │              │
       └──────────────┘            └──────────────┘
                (width, height)    (0, 0)        (width, 0)
```

```python
def pdfplumber_to_reportlab(y, page_height):
    return page_height - y
```

---

## Quick Reference

| 작업 | 도구 | 코드 |
|------|------|------|
| 텍스트 추출 | pdfplumber | `page.extract_text()` |
| 테이블 추출 | pdfplumber | `page.extract_tables()` |
| PDF 병합 | pypdf | `writer.add_page(page)` |
| PDF 분할 | pypdf | 페이지별 PdfWriter |
| PDF 생성 | reportlab | Canvas 또는 Platypus |
| 이미지 변환 | pdftoppm | `pdftoppm -jpeg -r 150` |
| 이미지 추출 | pdfimages | `pdfimages -all` |
| 폼 채우기 | - | `pdf-forms.md` 참조 |
| OCR | pytesseract | 이미지 변환 후 OCR |
| CLI 병합 | qpdf | `qpdf --empty --pages ...` |

---

## Troubleshooting

### 텍스트 추출 불가 (스캔 PDF)
- **해결**: OCR 사용 (pytesseract + pdf2image)

### 테이블 없음
- **해결**: `text` strategy 사용 (선 없는 테이블)

### "Index out of range"
- **원인**: 페이지 번호 오류 (0-indexed)
- **해결**: `reader.pages[0]`이 첫 페이지

### 한글 깨짐
- **해결**: UTF-8 인코딩 확인, 폰트 임베딩 확인

### 메모리 부족 (대용량)
- **해결**: 페이지별 처리, `read_only=True`, `page.flush_cache()`

### 손상된 PDF
- **해결**: `qpdf --check input.pdf` → `qpdf --replace-input`
