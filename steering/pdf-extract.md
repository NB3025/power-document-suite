# PDF 텍스트/테이블 추출 및 조작

PDF에서 텍스트, 테이블 추출 및 병합/분할 가이드입니다.

## Overview

PDF 파일에서 텍스트와 테이블을 추출하고, 여러 PDF를 병합/분할하는 방법을 다룹니다. pdfplumber와 pypdf 라이브러리를 사용합니다.

## Workflow Decision Tree

```
PDF 작업 유형?
├── 텍스트 추출 → pdfplumber (권장) 또는 pypdf
├── 테이블 추출 → pdfplumber + pandas
├── PDF 병합/분할 → pypdf
├── 워터마크 추가 → pypdf + reportlab
├── PDF → 이미지 → pdftoppm 또는 pdf2image
├── 폼 채우기 → pdf-forms.md
└── 메타데이터 수정 → pypdf
```

## Prerequisites

```bash
pip install pdfplumber pypdf pandas reportlab

# 이미지 변환용
brew install poppler  # macOS
# sudo apt-get install poppler-utils  # Ubuntu
```

---

## ⚠️ CRITICAL: 자주 발생하는 오류

### ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `reader.pages[1]` (첫 페이지) | `reader.pages[0]` | 0-indexed |
| pdfplumber 좌표 → reportlab 직접 사용 | `reportlab_y = page_height - pdfplumber_y` | 좌표 변환 필요 |
| `writer.write(output_path)` | `writer.write(file_obj)` 또는 `writer.write("path")` | 파일 객체/경로 |
| 닫지 않은 PdfWriter | `writer.close()` 호출 | 리소스 정리 |
| 테이블 없는 페이지에서 추출 | 먼저 테이블 존재 확인 | None 체크 |

---

## 텍스트 추출

### pdfplumber (권장)

레이아웃 기반 추출로 더 정확한 결과를 제공합니다.

```python
import pdfplumber

with pdfplumber.open("document.pdf") as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        if text:
            print(f"=== Page {i+1} ===")
            print(text)
        else:
            print(f"=== Page {i+1}: No text found ===")
```

### 특정 영역 추출

```python
import pdfplumber

with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]

    # 페이지 크기
    print(f"Width: {page.width}, Height: {page.height}")

    # 특정 영역만 추출 (bbox: x0, y0, x1, y1)
    # ⚠️ pdfplumber는 왼쪽 상단 기준!
    cropped = page.within_bbox((0, 0, page.width/2, page.height/2))  # 왼쪽 상단 1/4
    text = cropped.extract_text()
    print(text)
```

### pypdf

더 빠르지만 레이아웃 정보가 덜 정확합니다.

```python
from pypdf import PdfReader

reader = PdfReader("document.pdf")
print(f"Total pages: {len(reader.pages)}")

for i, page in enumerate(reader.pages):
    text = page.extract_text()
    print(f"=== Page {i+1} ===")
    print(text)
```

### 대용량 PDF 처리

```python
import pdfplumber

def extract_text_streaming(pdf_path, output_path):
    """대용량 PDF를 스트리밍 방식으로 처리합니다."""
    with pdfplumber.open(pdf_path) as pdf:
        with open(output_path, 'w', encoding='utf-8') as f:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    f.write(f"=== Page {i+1} ===\n")
                    f.write(text)
                    f.write("\n\n")
                # 메모리 절약
                page.flush_cache()

extract_text_streaming("large_document.pdf", "output.txt")
```

---

## 테이블 추출

### 기본 테이블 추출

```python
import pdfplumber
import pandas as pd

with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]

    # 모든 테이블 추출
    tables = page.extract_tables()

    if not tables:
        print("No tables found on this page")
    else:
        for i, table in enumerate(tables):
            if table and len(table) > 1:
                df = pd.DataFrame(table[1:], columns=table[0])
                print(f"=== Table {i+1} ===")
                print(df)
                df.to_csv(f"table_{i+1}.csv", index=False)
```

### 테이블 추출 옵션

```python
import pdfplumber

with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]

    # 선 기반 테이블 (기본)
    table_settings_lines = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
    }

    # 텍스트 기반 테이블 (선이 없는 경우)
    table_settings_text = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 3,
        "join_tolerance": 3,
    }

    # 명시적 좌표 (특정 영역만)
    table_settings_explicit = {
        "vertical_strategy": "explicit",
        "horizontal_strategy": "explicit",
        "explicit_vertical_lines": [100, 200, 300, 400],
        "explicit_horizontal_lines": [50, 100, 150, 200],
    }

    tables = page.extract_tables(table_settings_lines)
```

### 모든 페이지에서 테이블 추출

```python
import pdfplumber
import pandas as pd

def extract_all_tables(pdf_path):
    """모든 페이지에서 테이블을 추출합니다."""
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()

            for table_idx, table in enumerate(tables):
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df['source_page'] = page_num + 1
                    df['table_index'] = table_idx + 1
                    all_tables.append(df)

    if all_tables:
        combined = pd.concat(all_tables, ignore_index=True)
        return combined
    return None

df = extract_all_tables("report.pdf")
if df is not None:
    df.to_excel("all_tables.xlsx", index=False)
```

---

## PDF 병합

```python
from pypdf import PdfWriter

def merge_pdfs(pdf_files, output_path):
    """여러 PDF를 병합합니다."""
    writer = PdfWriter()

    for pdf_file in pdf_files:
        writer.append(pdf_file)

    writer.write(output_path)
    writer.close()

# 사용
merge_pdfs(["doc1.pdf", "doc2.pdf", "doc3.pdf"], "merged.pdf")
```

### 특정 페이지만 병합

```python
from pypdf import PdfReader, PdfWriter

def merge_specific_pages(inputs, output_path):
    """
    특정 페이지만 병합합니다.
    inputs: [(pdf_path, [page_numbers]), ...] - page_numbers는 0-indexed
    """
    writer = PdfWriter()

    for pdf_path, pages in inputs:
        reader = PdfReader(pdf_path)
        for page_num in pages:
            if 0 <= page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])

    writer.write(output_path)
    writer.close()

# 사용: doc1의 1,3페이지 + doc2의 2페이지 병합
merge_specific_pages([
    ("doc1.pdf", [0, 2]),  # 1페이지, 3페이지
    ("doc2.pdf", [1]),     # 2페이지
], "selected_pages.pdf")
```

---

## PDF 분할

### 각 페이지를 별도 파일로

```python
from pypdf import PdfReader, PdfWriter

def split_pdf_to_pages(input_path, output_prefix):
    """각 페이지를 별도 PDF로 분할합니다."""
    reader = PdfReader(input_path)

    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        writer.write(f"{output_prefix}_page_{i+1}.pdf")
        writer.close()

split_pdf_to_pages("document.pdf", "output")
```

### 페이지 범위로 분할

```python
from pypdf import PdfReader, PdfWriter

def extract_pages(input_path, output_path, start_page, end_page):
    """
    특정 페이지 범위를 추출합니다.
    start_page, end_page: 1-indexed (사용자 친화적)
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    # 1-indexed를 0-indexed로 변환
    for i in range(start_page - 1, min(end_page, len(reader.pages))):
        writer.add_page(reader.pages[i])

    writer.write(output_path)
    writer.close()

# 3-5페이지 추출
extract_pages("document.pdf", "pages_3_to_5.pdf", 3, 5)
```

---

## 워터마크 추가

```python
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO

def create_watermark(text, opacity=0.3):
    """워터마크 PDF를 생성합니다."""
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    c.saveState()
    c.setFont("Helvetica", 50)
    c.setFillColorRGB(0.5, 0.5, 0.5, opacity)
    c.translate(letter[0]/2, letter[1]/2)
    c.rotate(45)
    c.drawCentredString(0, 0, text)
    c.restoreState()

    c.save()
    buffer.seek(0)
    return buffer

def add_watermark(input_path, output_path, watermark_text):
    """PDF에 워터마크를 추가합니다."""
    reader = PdfReader(input_path)
    watermark_reader = PdfReader(create_watermark(watermark_text))
    writer = PdfWriter()

    for page in reader.pages:
        page.merge_page(watermark_reader.pages[0])
        writer.add_page(page)

    writer.write(output_path)
    writer.close()

add_watermark("document.pdf", "watermarked.pdf", "CONFIDENTIAL")
```

---

## PDF → 이미지 변환

### pdftoppm (권장)

```bash
# 모든 페이지를 JPEG로
pdftoppm -jpeg -r 150 document.pdf output

# 특정 페이지만 (1-3페이지)
pdftoppm -jpeg -r 300 -f 1 -l 3 document.pdf pages

# PNG로 변환
pdftoppm -png -r 200 document.pdf output
```

### Python (pdf2image)

```python
from pdf2image import convert_from_path

def pdf_to_images(pdf_path, output_dir, dpi=150, fmt="jpeg"):
    """PDF를 이미지로 변환합니다."""
    images = convert_from_path(pdf_path, dpi=dpi)

    for i, img in enumerate(images):
        output_path = f"{output_dir}/page_{i+1}.{fmt}"
        img.save(output_path, fmt.upper())
        print(f"Saved: {output_path}")

pdf_to_images("document.pdf", "images", dpi=150)
```

---

## 메타데이터 조회/수정

### 메타데이터 조회

```python
from pypdf import PdfReader

reader = PdfReader("document.pdf")

# 메타데이터 조회
print("=== Metadata ===")
if reader.metadata:
    for key, value in reader.metadata.items():
        print(f"{key}: {value}")

print(f"\nTotal pages: {len(reader.pages)}")
```

### 메타데이터 수정

```python
from pypdf import PdfReader, PdfWriter

def update_metadata(input_path, output_path, metadata):
    """PDF 메타데이터를 수정합니다."""
    reader = PdfReader(input_path)
    writer = PdfWriter()

    writer.append(reader)
    writer.add_metadata(metadata)

    writer.write(output_path)
    writer.close()

# 사용
update_metadata("document.pdf", "updated.pdf", {
    "/Title": "Updated Title",
    "/Author": "Author Name",
    "/Subject": "Document Subject",
    "/Keywords": "keyword1, keyword2",
    "/Creator": "My Application"
})
```

---

## 페이지 회전

```python
from pypdf import PdfReader, PdfWriter

def rotate_pages(input_path, output_path, rotation, pages=None):
    """
    페이지를 회전합니다.
    rotation: 90, 180, 270
    pages: None (모든 페이지) 또는 [0, 2, 4] (특정 페이지, 0-indexed)
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        if pages is None or i in pages:
            page.rotate(rotation)
        writer.add_page(page)

    writer.write(output_path)
    writer.close()

# 모든 페이지 90도 회전
rotate_pages("document.pdf", "rotated.pdf", 90)

# 1, 3페이지만 180도 회전
rotate_pages("document.pdf", "rotated.pdf", 180, pages=[0, 2])
```

---

## 문자 위치 정보 추출

```python
import pdfplumber

with pdfplumber.open("document.pdf") as pdf:
    page = pdf.pages[0]

    # 문자 위치 정보
    for char in page.chars[:20]:  # 처음 20자
        print(f"'{char['text']}' at ({char['x0']:.1f}, {char['y0']:.1f}) - ({char['x1']:.1f}, {char['y1']:.1f})")

    # 특정 텍스트 찾기
    def find_text_position(page, search_text):
        """텍스트의 위치를 찾습니다."""
        text = page.extract_text()
        if search_text in text:
            # 단어 단위로 검색
            for word in page.extract_words():
                if search_text in word['text']:
                    return {
                        'text': word['text'],
                        'x0': word['x0'],
                        'y0': word['top'],  # pdfplumber의 top
                        'x1': word['x1'],
                        'y1': word['bottom']
                    }
        return None

    position = find_text_position(page, "Invoice")
    if position:
        print(f"Found '{position['text']}' at ({position['x0']:.1f}, {position['y0']:.1f})")
```

---

## 좌표 체계 비교

```
pdfplumber 좌표 (왼쪽 상단 기준)         reportlab 좌표 (왼쪽 하단 기준)
┌─────────────────┐ (0, 0)               ┌─────────────────┐ (0, height)
│                 │                       │                 │
│                 │                       │                 │
│                 │                       │                 │
└─────────────────┘ (width, height)      └─────────────────┘ (width, 0)
                                          (0, 0)
```

### 좌표 변환

```python
def pdfplumber_to_reportlab(y, page_height):
    """pdfplumber 좌표를 reportlab 좌표로 변환합니다."""
    return page_height - y

def reportlab_to_pdfplumber(y, page_height):
    """reportlab 좌표를 pdfplumber 좌표로 변환합니다."""
    return page_height - y
```

---

## Troubleshooting

### "Cannot extract text" (텍스트 추출 불가)
- **원인**: 스캔된 PDF (이미지 기반)
- **해결**: OCR 사용 (pytesseract + pdf2image)

```python
from pdf2image import convert_from_path
import pytesseract

images = convert_from_path("scanned.pdf")
for i, img in enumerate(images):
    text = pytesseract.image_to_string(img, lang='eng')
    print(f"Page {i+1}: {text}")
```

### "No tables found" (테이블 없음)
- **원인**: 선이 없는 테이블
- **해결**: `text` strategy 사용

### "Index out of range"
- **원인**: 페이지 번호 오류 (1-indexed vs 0-indexed)
- **해결**: `reader.pages[0]`이 첫 페이지

### 한글 깨짐
- **원인**: 인코딩 문제
- **해결**: UTF-8 인코딩 확인, 폰트 임베딩 확인

### 메모리 부족 (대용량 PDF)
- **원인**: 전체 PDF 로드
- **해결**: 페이지별 처리, `page.flush_cache()` 사용

---

## Quick Reference

### pdfplumber 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `pdf.pages` | 페이지 리스트 |
| `page.extract_text()` | 텍스트 추출 |
| `page.extract_tables()` | 테이블 추출 |
| `page.chars` | 문자 위치 정보 |
| `page.extract_words()` | 단어 위치 정보 |
| `page.within_bbox((x0, y0, x1, y1))` | 영역 크롭 |

### pypdf 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `PdfReader(path)` | PDF 읽기 |
| `PdfWriter()` | PDF 쓰기 |
| `writer.append(reader)` | PDF 추가 |
| `writer.add_page(page)` | 페이지 추가 |
| `page.rotate(degrees)` | 페이지 회전 |
| `page.merge_page(other)` | 페이지 병합 |
| `writer.add_metadata({})` | 메타데이터 추가 |
