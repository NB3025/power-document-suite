# PDF 폼 작성

PDF 폼 필드 작성 및 비-Fillable PDF에 텍스트 오버레이 가이드입니다.

## Overview

PDF 폼은 두 가지 유형이 있습니다:
1. **Fillable PDF**: 대화형 폼 필드가 내장된 PDF (텍스트 필드, 체크박스, 라디오 버튼 등)
2. **Non-Fillable PDF**: 폼 필드 없이 시각적 양식만 있는 PDF (스캔 문서, 이미지 기반 양식)

각 유형에 따라 다른 접근 방식이 필요합니다.

## Workflow Decision Tree

```
PDF 폼 작성?
├── 1. 폼 필드 존재 여부 확인 → check_fillable_fields()
├── Fillable (폼 필드 있음)
│   ├── 필드 정보 추출 → get_form_fields()
│   ├── 필드 값 설정 → update_page_form_field_values()
│   └── 저장
└── Non-Fillable (폼 필드 없음)
    ├── 구조 추출 시도 → analyze_form_layout() (pdfplumber)
    ├── 구조 있음 → Approach A: 구조 기반 좌표
    ├── 구조 없음 → Approach B: 시각 추정 (이미지 분석)
    ├── 부분 구조 → Hybrid: 구조 + 시각
    ├── 바운딩 박스 검증 → check_bounding_boxes()
    ├── 텍스트 오버레이 생성 (reportlab)
    └── 페이지 병합 → 저장
```

## Prerequisites

```bash
pip install pypdf pdfplumber reportlab pdf2image Pillow
brew install poppler imagemagick  # macOS
```

---

## ⚠️ CRITICAL: 자주 발생하는 오류

### ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `"/Yes"` (모든 체크박스) | 실제 checked_value 사용 | 체크박스마다 값이 다름 |
| reportlab에 pdfplumber 좌표 직접 사용 | `page_height - y` 변환 | 좌표 체계가 다름 |
| `writer.write(output_path)` | `writer.write(file_obj)` | 파일 객체로 쓰기 |
| 폼 필드 없는 PDF에 `update_page_form_field_values` | 텍스트 오버레이 사용 | Non-Fillable 처리 |
| 대략적인 좌표로 바로 채우기 | 줌 보정 후 채우기 | 텍스트 위치 부정확 |
| 바운딩 박스 검증 없이 채우기 | `check_bounding_boxes()` 실행 | 겹침/크기 오류 방지 |
| 바운딩 박스가 레이블과 겹침 | 레이블/입력 영역 분리 | 텍스트 덮어쓰기 방지 |
| 텍스트 크기 미고려 | 바운딩 박스 높이 충분히 확보 | 텍스트 잘림 방지 |

---

## Fillable PDF 폼 작성

### Step 1: 폼 필드 존재 확인

```python
from pypdf import PdfReader

def check_fillable_fields(pdf_path):
    """PDF에 폼 필드가 있는지 확인합니다."""
    reader = PdfReader(pdf_path)
    fields = reader.get_fields()

    if not fields:
        print("⚠️ 폼 필드가 없습니다. Non-Fillable 방식을 사용하세요.")
        return False

    print(f"✅ {len(fields)}개의 폼 필드를 발견했습니다.")
    return True

check_fillable_fields("form.pdf")
```

Fillable이면 아래 진행. 아니면 "Non-Fillable PDF 폼 작성" 섹션으로 이동.

### Step 2: 폼 필드 정보 추출

```python
from pypdf import PdfReader

def get_form_fields(pdf_path):
    """PDF 폼 필드의 상세 정보를 조회합니다."""
    reader = PdfReader(pdf_path)
    fields = reader.get_fields()

    if not fields:
        print("폼 필드가 없습니다.")
        return {}

    field_info = []
    for name, field in fields.items():
        info = {
            "field_id": name,
            "type": field.get("/FT", "Unknown"),
            "value": field.get("/V", ""),
            "default": field.get("/DV", ""),
        }

        # 체크박스/라디오 옵션 확인
        if "/Opt" in field:
            info["options"] = field["/Opt"]

        # 필드 플래그 (읽기 전용 등)
        if "/Ff" in field:
            info["flags"] = field["/Ff"]

        field_info.append(info)

        print(f"Field: {name}")
        print(f"  Type: {info['type']}")
        print(f"  Current Value: {info['value']}")
        if "options" in info:
            print(f"  Options: {info['options']}")
        print()

    return field_info

fields = get_form_fields("form.pdf")
```

### 필드 유형 (Field Types)

| 유형 코드 | 설명 | 값 형식 |
|-----------|------|---------|
| `/Tx` | 텍스트 필드 | 문자열 |
| `/Btn` | 버튼 (체크박스/라디오) | `/Yes`, `/Off`, `/Option1` 등 |
| `/Ch` | 선택 (드롭다운/리스트) | 옵션 값 문자열 |
| `/Sig` | 서명 필드 | 서명 데이터 |

### Step 3: 폼 필드 채우기

```python
from pypdf import PdfReader, PdfWriter

def fill_pdf_form(input_path, output_path, field_data):
    """PDF 폼 필드를 채웁니다."""
    reader = PdfReader(input_path)
    writer = PdfWriter()

    # 모든 페이지 추가
    writer.append(reader)

    # 첫 페이지 필드 업데이트 (단일 페이지 폼)
    writer.update_page_form_field_values(
        writer.pages[0],
        field_data
    )

    # 저장
    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"✅ 저장 완료: {output_path}")

# 사용 예시
field_data = {
    "name": "홍길동",
    "email": "hong@example.com",
    "date": "2025-01-27",
    "agree_checkbox": "/Yes"  # ⚠️ 실제 checked_value 확인 필요
}

fill_pdf_form("form.pdf", "filled_form.pdf", field_data)
```

### 여러 페이지 폼 채우기

```python
from pypdf import PdfReader, PdfWriter

def fill_multipage_form(input_path, output_path, page_field_data):
    """
    여러 페이지 폼을 채웁니다.
    page_field_data: {page_index: {field_name: value, ...}, ...}
    ⚠️ page_index는 0-based입니다.
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.append(reader)

    for page_num, field_data in page_field_data.items():
        if page_num < len(writer.pages):
            writer.update_page_form_field_values(
                writer.pages[page_num],
                field_data
            )
        else:
            print(f"⚠️ 페이지 {page_num}이 존재하지 않습니다.")

    with open(output_path, "wb") as f:
        writer.write(f)

# 사용 예시
page_field_data = {
    0: {"name": "홍길동", "date": "2025-01-27"},
    1: {"address": "서울시 강남구", "phone": "010-1234-5678"},
    2: {"signature_date": "2025-01-27"}
}

fill_multipage_form("multi_page_form.pdf", "filled.pdf", page_field_data)
```

---

## 체크박스 / 라디오 버튼 처리

### ⚠️ CRITICAL: checked_value 확인 필수

체크박스와 라디오 버튼의 값은 PDF마다 다릅니다. 반드시 실제 값을 확인하세요.

```python
from pypdf import PdfReader

def get_checkbox_values(pdf_path):
    """체크박스/라디오 버튼의 가능한 값을 확인합니다."""
    reader = PdfReader(pdf_path)

    for page in reader.pages:
        if "/Annots" in page:
            for annot in page["/Annots"]:
                annot_obj = annot.get_object()
                if annot_obj.get("/FT") == "/Btn":
                    field_name = annot_obj.get("/T", "Unknown")

                    # 가능한 값 확인
                    if "/AP" in annot_obj:
                        appearances = annot_obj["/AP"]
                        if "/N" in appearances:
                            possible_values = list(appearances["/N"].keys())
                            print(f"Field: {field_name}")
                            print(f"  Possible values: {possible_values}")
                            print(f"  (Use one of these as the value)")
                            print()

get_checkbox_values("form.pdf")
```

### 체크박스 값 설정

```python
# 일반적인 체크박스 값
field_data = {
    # 체크된 상태 (PDF마다 다를 수 있음)
    "checkbox1": "/Yes",      # 가장 일반적
    "checkbox2": "/On",       # 일부 PDF
    "checkbox3": "/1",        # 숫자 형식

    # 체크 해제
    "checkbox4": "/Off",      # 일반적
    "checkbox5": "/No",       # 일부 PDF
    "checkbox6": "/0",        # 숫자 형식
}
```

### 라디오 버튼 그룹 설정

```python
# 라디오 그룹에서 하나만 선택
field_data = {
    "gender_group": "/Male",     # 옵션 중 하나 선택
    "payment_method": "/Credit", # 다른 그룹
}
```

---

## Non-Fillable PDF 폼 작성

폼 필드가 없는 PDF에는 텍스트 오버레이를 사용합니다. 먼저 PDF 구조에서 좌표를 추출하고(더 정확), 실패하면 시각 추정으로 대체합니다.

### 3가지 접근법 선택

```
구조 추출 시도 (analyze_form_layout)
├── 텍스트 레이블 있음 → Approach A: 구조 기반 좌표 (권장)
├── 레이블 없음 (스캔/이미지) → Approach B: 시각 추정 (이미지 분석)
└── 일부 레이블만 있음 → Hybrid: 구조 + 시각
```

### Step 1: 구조 추출 시도

```python
import pdfplumber

def analyze_form_layout(pdf_path):
    """PDF 레이아웃을 분석하여 텍스트 위치를 확인합니다."""
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"=== Page {page_num + 1} ===")
            print(f"Size: {page.width} x {page.height}")
            print()

            # 텍스트 위치 (레이블 찾기용)
            words = page.extract_words()
            for word in words[:30]:  # 처음 30개
                print(f"'{word['text']}' at ({word['x0']:.1f}, {word['top']:.1f}) - ({word['x1']:.1f}, {word['bottom']:.1f})")

            # 선 찾기 (입력 영역 힌트)
            lines = page.lines
            if lines:
                print(f"\n{len(lines)} lines found")
                for line in lines[:10]:
                    print(f"Line: ({line['x0']:.1f}, {line['top']:.1f}) - ({line['x1']:.1f}, {line['bottom']:.1f})")

analyze_form_layout("form.pdf")
```

**결과 확인**: 의미 있는 레이블이 출력되면 **Approach A**. 스캔/이미지 기반이라 레이블이 없거나 `(cid:X)` 패턴만 나오면 **Approach B**. 일부만 있으면 **Hybrid**.

구조 추출의 JSON 출력에 포함되는 정보:
- **labels**: 모든 텍스트 요소와 정확한 좌표 (x0, top, x1, bottom -- PDF 포인트)
- **lines**: 행 경계를 정의하는 수평선
- **checkboxes**: 체크박스인 작은 사각형 (중심 좌표 포함)
- **row_boundaries**: 수평선에서 계산된 행 상단/하단 위치

### Step 2: PDF를 이미지로 변환 (위치 분석용)

```python
from pdf2image import convert_from_path

def pdf_to_images(pdf_path, output_dir, dpi=150):
    """PDF를 이미지로 변환하여 위치 분석에 사용합니다."""
    images = convert_from_path(pdf_path, dpi=dpi)

    for i, img in enumerate(images):
        output_path = f"{output_dir}/page_{i+1}.png"
        img.save(output_path, "PNG")
        print(f"Page {i+1}: {img.width}x{img.height} -> {output_path}")

    return images

# 이미지 생성 후 좌표 확인
pdf_to_images("form.pdf", "output_images")
```

---

### Approach A: 구조 기반 좌표 (권장)

`analyze_form_layout()`가 텍스트 레이블을 찾은 경우 사용.

#### A.1: 구조 분석

pdfplumber 출력을 읽고 식별:

1. **레이블 그룹**: 하나의 레이블을 구성하는 인접 텍스트 요소 (예: "Last" + "Name")
2. **행 구조**: 유사한 `top` 값을 가진 레이블은 같은 행
3. **필드 컬럼**: 입력 영역은 레이블 끝 이후 시작 (x0 = label.x1 + gap)
4. **체크박스**: 구조에서 직접 체크박스 좌표 사용

**좌표 체계**: PDF 좌표 -- y=0이 페이지 상단 (pdfplumber 기준), y가 아래로 증가.

#### A.2: 누락 요소 확인

구조 추출이 모든 요소를 감지하지 못할 수 있음:
- **원형 체크박스**: 사각형만 감지됨
- **복잡한 그래픽**: 장식 요소, 비표준 폼 컨트롤
- **흐리거나 밝은 색 요소**: 추출되지 않을 수 있음

이미지에는 보이지만 구조 추출에 없는 필드는 **Hybrid Approach** 참조.

#### A.3: fields.json 작성 (PDF 좌표)

각 필드의 입력 좌표를 구조에서 계산:

**텍스트 필드:**
- entry x0 = label x1 + 5 (레이블 뒤 작은 간격)
- entry x1 = 다음 레이블의 x0, 또는 행 경계
- entry top = 레이블 top과 동일
- entry bottom = 아래쪽 행 경계선, 또는 label bottom + row_height

**체크박스:**
- 구조 추출의 체크박스 사각형 좌표를 직접 사용

```json
{
  "pages": [
    {"page_number": 1, "pdf_width": 612, "pdf_height": 792}
  ],
  "form_fields": [
    {
      "page_number": 1,
      "description": "성 입력 필드",
      "field_label": "Last Name",
      "label_bounding_box": [43, 63, 87, 73],
      "entry_bounding_box": [92, 63, 260, 79],
      "entry_text": {"text": "홍", "font_size": 10}
    },
    {
      "page_number": 1,
      "description": "시민권 Yes 체크박스",
      "field_label": "Yes",
      "label_bounding_box": [260, 200, 280, 210],
      "entry_bounding_box": [285, 197, 292, 205],
      "entry_text": {"text": "X"}
    }
  ]
}
```

**중요**: `pdf_width`/`pdf_height` 사용, pdfplumber 구조의 좌표를 직접 사용.

#### A.4: 바운딩 박스 검증

```python
check_bounding_boxes(fields_data)  # 아래 "바운딩 박스 검증" 섹션 참조
```

겹치는 바운딩 박스와 폰트 크기에 비해 너무 작은 입력 박스를 검사. 에러 수정 후 채우기 진행.

---

### Approach B: 시각 추정 (대체)

스캔/이미지 기반 PDF로 구조 추출에서 텍스트 레이블이 없는 경우 사용 (예: 모든 텍스트가 "(cid:X)" 패턴).

#### B.1: 이미지 변환

```python
pdf_to_images("input.pdf", "images/")  # 위의 pdf_to_images() 함수 사용
```

#### B.2: 초기 필드 식별

각 페이지 이미지를 분석하여 **대략적인** 필드 위치 추정:
- 폼 필드 레이블과 대략적 위치
- 입력 영역 (선, 박스, 빈 공간)
- 체크박스와 대략적 위치

각 필드의 대략적 픽셀 좌표를 기록 (정밀할 필요 없음).

#### B.3: 줌 보정 (정확도를 위해 CRITICAL)

각 필드에 대해 추정 위치 주변을 잘라 좌표를 정밀하게 보정.

**ImageMagick으로 줌 크롭 생성:**

```bash
magick <page_image> -crop <width>x<height>+<x>+<y> +repage <crop_output.png>
```

- `<x>, <y>` = 크롭 영역의 좌상단 (대략 추정값 - 패딩)
- `<width>, <height>` = 크롭 영역 크기 (필드 영역 + ~50px 패딩)

**예시:** "Name" 필드가 (100, 150) 근처로 추정될 때:

```bash
magick images/page_1.png -crop 300x80+50+120 +repage crops/name_field.png
```

(`magick` 없으면 `convert`로 동일 인수 사용)

**크롭 이미지 분석:**
1. 입력 영역 시작 정확한 픽셀 식별 (레이블 뒤)
2. 입력 영역 끝 식별 (다음 필드 또는 가장자리 전)
3. 입력 선/박스의 상단과 하단 식별

**크롭 좌표를 전체 이미지 좌표로 변환:**

```
full_x = crop_x + crop_offset_x
full_y = crop_y + crop_offset_y
```

예시: 크롭이 (50, 120)에서 시작, 크롭 내 입력 박스가 (52, 18)에서 시작:
- entry_x0 = 52 + 50 = 102
- entry_top = 18 + 120 = 138

**각 필드에 반복**, 인접 필드는 하나의 크롭으로 묶기.

#### B.4: fields.json 작성 (이미지 좌표)

```json
{
  "pages": [
    {"page_number": 1, "image_width": 1700, "image_height": 2200}
  ],
  "form_fields": [
    {
      "page_number": 1,
      "description": "성 입력 필드",
      "field_label": "Last Name",
      "label_bounding_box": [120, 175, 242, 198],
      "entry_bounding_box": [255, 175, 720, 218],
      "entry_text": {"text": "홍", "font_size": 10}
    }
  ]
}
```

**중요**: `image_width`/`image_height` 사용, 줌 분석에서 보정된 픽셀 좌표 사용.

#### B.5: 바운딩 박스 검증

```python
check_bounding_boxes(fields_data)  # 아래 "바운딩 박스 검증" 섹션 참조
```

---

### Hybrid Approach: 구조 + 시각

구조 추출이 대부분 필드에서 작동하지만 일부 요소를 놓친 경우 사용.

1. **Approach A** -- 구조에서 감지된 필드에 사용
2. **이미지 변환** -- 누락 필드의 시각 분석용
3. **줌 보정** (Approach B) -- 누락 필드에 사용
4. **좌표 통합**: 구조 추출 필드는 `pdf_width`/`pdf_height` 사용. 시각 추정 필드는 이미지 -> PDF 좌표 변환:
   - `pdf_x = image_x * (pdf_width / image_width)`
   - `pdf_y = image_y * (pdf_height / image_height)`
5. **단일 좌표 체계 사용** -- fields.json에서 모든 좌표를 PDF 좌표(`pdf_width`/`pdf_height`)로 통일

---

## 바운딩 박스 패턴

```
패턴 1: 레이블 + 밑줄
  Name: ____________________
  └────┘ └─────────────────┘
  레이블    입력 영역

패턴 2: 밑줄 아래 레이블
  _________________________
  Name
  └────────────────────────┘ 입력 영역
  └────┘ 레이블

패턴 3: 박스 안 레이블
  ┌────────────────────────┐
  │ Name:                  │
  └────────────────────────┘
  └─────┘ └───────────────┘
  레이블     입력 영역

패턴 4: 체크박스
  □ Yes    □ No
  └┘       └┘  <- 체크박스 영역 (작게)
    └──┘    └─┘ <- 레이블 영역
```

---

## 바운딩 박스 검증

### 검증 체크리스트

- [ ] 레이블과 입력 영역 바운딩 박스가 겹치지 않음
- [ ] 입력 영역이 텍스트 높이보다 충분히 큼 (font_size + 2~4pt)
- [ ] 체크박스는 박스 영역만 포함 (레이블 제외)
- [ ] 좌표가 페이지 범위 내에 있음

### 바운딩 박스 검증 함수

```python
def check_bbox_intersection(bbox1, bbox2):
    """두 바운딩 박스가 겹치는지 확인합니다."""
    # bbox: [left, top, right, bottom]
    left1, top1, right1, bottom1 = bbox1
    left2, top2, right2, bottom2 = bbox2

    # 겹치지 않는 조건
    if right1 < left2 or right2 < left1:
        return False
    if bottom1 < top2 or bottom2 < top1:
        return False

    return True

def validate_form_fields(form_fields):
    """폼 필드 정의를 검증합니다."""
    errors = []

    for field in form_fields["form_fields"]:
        label_bbox = field.get("label_bounding_box")
        entry_bbox = field.get("entry_bounding_box")

        if label_bbox and entry_bbox:
            if check_bbox_intersection(label_bbox, entry_bbox):
                errors.append(f"❌ '{field['field_label']}': 레이블과 입력 영역이 겹칩니다.")

        # 입력 영역 높이 확인
        if entry_bbox:
            height = entry_bbox[3] - entry_bbox[1]
            font_size = field.get("entry_text", {}).get("font_size", 14)
            if height < font_size:
                errors.append(f"❌ '{field['field_label']}': 입력 영역 높이({height})가 폰트 크기({font_size})보다 작습니다.")

    if errors:
        print("검증 실패:")
        for error in errors:
            print(f"  {error}")
        return False

    print("✅ 모든 검증 통과")
    return True

def check_bounding_boxes(fields_data):
    """
    fields.json 데이터의 바운딩 박스를 종합 검증합니다.
    좌표 체계(PDF/이미지)에 관계없이 겹침, 크기, 범위를 검사합니다.
    """
    errors = []

    # 페이지 크기 정보 수집
    page_sizes = {}
    for page_info in fields_data.get("pages", []):
        pn = page_info["page_number"]
        # PDF 좌표 또는 이미지 좌표 중 사용된 것을 감지
        if "pdf_width" in page_info:
            page_sizes[pn] = (page_info["pdf_width"], page_info["pdf_height"])
        elif "image_width" in page_info:
            page_sizes[pn] = (page_info["image_width"], page_info["image_height"])

    for field in fields_data.get("form_fields", []):
        label = field.get("field_label", "Unknown")
        label_bbox = field.get("label_bounding_box")
        entry_bbox = field.get("entry_bounding_box")
        page_num = field.get("page_number", 1)

        # 겹침 검사
        if label_bbox and entry_bbox:
            if check_bbox_intersection(label_bbox, entry_bbox):
                errors.append(f"겹침: '{label}' 레이블과 입력 영역이 겹칩니다.")

        # 입력 영역 높이 vs 폰트 크기
        if entry_bbox:
            height = entry_bbox[3] - entry_bbox[1]
            font_size = field.get("entry_text", {}).get("font_size", 14)
            if height < font_size:
                errors.append(
                    f"크기: '{label}' 입력 높이({height:.1f})가 "
                    f"폰트 크기({font_size})보다 작습니다."
                )

        # 페이지 범위 검사
        if page_num in page_sizes and entry_bbox:
            pw, ph = page_sizes[page_num]
            left, top, right, bottom = entry_bbox
            if right > pw or bottom > ph or left < 0 or top < 0:
                errors.append(
                    f"범위: '{label}' 입력 영역이 페이지 범위를 벗어납니다. "
                    f"bbox={entry_bbox}, page=({pw}, {ph})"
                )

    if errors:
        print(f"❌ {len(errors)}개 오류 발견:")
        for e in errors:
            print(f"  - {e}")
        return False

    print("✅ 모든 바운딩 박스 검증 통과")
    return True
```

---

## 텍스트 오버레이 생성 및 병합

### 텍스트 오버레이 생성

```python
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO

def create_text_overlay(page_size, text_items):
    """
    텍스트 오버레이 PDF를 생성합니다.
    text_items: [(x, y, text, font_size, font_color), ...]
    ⚠️ 좌표는 reportlab 기준 (왼쪽 하단 원점)
    """
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=page_size)

    for item in text_items:
        x, y, text, font_size = item[:4]
        font_color = item[4] if len(item) > 4 else "000000"

        # 색상 설정 (RRGGBB)
        r = int(font_color[0:2], 16) / 255
        g = int(font_color[2:4], 16) / 255
        b = int(font_color[4:6], 16) / 255
        c.setFillColorRGB(r, g, b)

        c.setFont("Helvetica", font_size)
        c.drawString(x, y, text)

    c.save()
    buffer.seek(0)
    return buffer

def fill_pdf_with_overlay(input_path, output_path, page_text_items):
    """
    PDF에 텍스트 오버레이를 추가합니다.
    page_text_items: {page_index: [(x, y, text, font_size), ...], ...}
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        if i in page_text_items and page_text_items[i]:
            # 페이지 크기 가져오기
            media_box = page.mediabox
            page_width = float(media_box.width)
            page_height = float(media_box.height)

            # 오버레이 생성
            overlay_buffer = create_text_overlay(
                (page_width, page_height),
                page_text_items[i]
            )
            overlay_reader = PdfReader(overlay_buffer)

            # 병합
            page.merge_page(overlay_reader.pages[0])

        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"✅ 저장 완료: {output_path}")

# 사용 예시
# ⚠️ 좌표는 reportlab 기준 (왼쪽 하단 원점)
page_text_items = {
    0: [  # 첫 번째 페이지 (0-indexed)
        (100, 700, "홍길동", 12),
        (100, 680, "hong@example.com", 12),
        (400, 700, "2025-01-27", 12),
        (72, 300, "X", 14),  # 체크박스
    ],
    1: [  # 두 번째 페이지
        (100, 700, "서울시 강남구", 12),
    ]
}

fill_pdf_with_overlay("form.pdf", "filled.pdf", page_text_items)
```

---

## 좌표 체계 비교 및 변환

### 좌표 체계

```
pdfplumber (왼쪽 상단 원점)     reportlab (왼쪽 하단 원점)
(0, 0) ┌──────────────┐        ┌──────────────┐ (0, height)
       │              │        │              │
       │              │        │              │
       │              │        │              │
       └──────────────┘        └──────────────┘
                (width, height) (0, 0)        (width, 0)
```

### 좌표 변환 함수

```python
def pdfplumber_to_reportlab(y, page_height):
    """pdfplumber 좌표를 reportlab 좌표로 변환합니다."""
    return page_height - y

def reportlab_to_pdfplumber(y, page_height):
    """reportlab 좌표를 pdfplumber 좌표로 변환합니다."""
    return page_height - y

def image_to_pdf_coords(img_x, img_y, img_width, img_height, pdf_width, pdf_height):
    """
    이미지 좌표를 PDF 좌표로 변환합니다 (pdfplumber 기준).
    이미지와 pdfplumber 모두 왼쪽 상단 원점이므로 스케일만 적용.
    """
    scale_x = pdf_width / img_width
    scale_y = pdf_height / img_height

    pdf_x = img_x * scale_x
    pdf_y = img_y * scale_y

    return pdf_x, pdf_y

def image_to_reportlab_coords(img_x, img_y, img_width, img_height, pdf_width, pdf_height):
    """
    이미지 좌표를 reportlab 좌표로 변환합니다.
    이미지는 왼쪽 상단 기준, reportlab은 왼쪽 하단 기준이므로 y 반전.
    """
    scale_x = pdf_width / img_width
    scale_y = pdf_height / img_height

    pdf_x = img_x * scale_x
    # 이미지는 왼쪽 상단 기준, PDF(reportlab)는 왼쪽 하단 기준
    pdf_y = pdf_height - (img_y * scale_y)

    return pdf_x, pdf_y
```

### fields.json 좌표 체계 선택

| 키 | 좌표 체계 | 사용 시점 |
|----|-----------|----------|
| `pdf_width` / `pdf_height` | PDF 포인트 | 구조 추출 (Approach A) |
| `image_width` / `image_height` | 이미지 픽셀 | 시각 추정 (Approach B) |

### 완전한 좌표 변환 예제

```python
import pdfplumber

def find_text_and_convert_coords(pdf_path, search_text):
    """텍스트를 찾고 reportlab 좌표로 변환합니다."""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        page_height = page.height

        for word in page.extract_words():
            if search_text.lower() in word['text'].lower():
                # pdfplumber 좌표 (왼쪽 상단 기준)
                plumber_x = word['x0']
                plumber_y = word['top']

                # reportlab 좌표로 변환 (왼쪽 하단 기준)
                reportlab_x = plumber_x
                reportlab_y = page_height - plumber_y

                print(f"Found '{word['text']}':")
                print(f"  pdfplumber: ({plumber_x:.1f}, {plumber_y:.1f})")
                print(f"  reportlab:  ({reportlab_x:.1f}, {reportlab_y:.1f})")

                return reportlab_x, reportlab_y

    return None

# "Name:" 레이블 오른쪽에 텍스트 추가할 위치 찾기
coords = find_text_and_convert_coords("form.pdf", "Name:")
```

---

## 한글 폰트 사용

```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def register_korean_font():
    """한글 폰트를 등록합니다."""
    # macOS
    try:
        pdfmetrics.registerFont(TTFont('AppleGothic', '/System/Library/Fonts/AppleGothic.ttf'))
        return 'AppleGothic'
    except:
        pass

    # Windows
    try:
        pdfmetrics.registerFont(TTFont('Malgun', 'C:/Windows/Fonts/malgun.ttf'))
        return 'Malgun'
    except:
        pass

    # Linux (Noto Sans CJK)
    try:
        pdfmetrics.registerFont(TTFont('NotoSansCJK', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'))
        return 'NotoSansCJK'
    except:
        pass

    print("⚠️ 한글 폰트를 찾을 수 없습니다.")
    return 'Helvetica'

# 사용
font_name = register_korean_font()
c.setFont(font_name, 12)
c.drawString(100, 700, "한글 텍스트")
```

---

## Troubleshooting

### "폼 필드가 없습니다"
- **원인**: PDF에 대화형 폼 필드가 없음 (스캔 문서, 이미지 기반)
- **해결**: Non-Fillable 방식으로 텍스트 오버레이 사용

### 체크박스가 체크되지 않음
- **원인**: 잘못된 checked_value 사용
- **해결**: `get_checkbox_values()`로 실제 값 확인 후 사용

### 텍스트가 잘못된 위치에 표시됨
- **원인**: 좌표 체계 혼동 (pdfplumber vs reportlab) 또는 좌표 보정 없이 대략적 좌표 사용
- **해결**: `pdfplumber_to_reportlab()` 변환 사용, 줌 보정 (Approach B.3) 수행

### 한글이 깨짐
- **원인**: 한글 폰트 미등록
- **해결**: `register_korean_font()`로 한글 폰트 등록 후 사용

### "Index out of range" 페이지 오류
- **원인**: 페이지 인덱스 오류 (0-indexed)
- **해결**: `pages[0]`이 첫 페이지, `pages[1]`이 두 번째 페이지

### 오버레이 텍스트가 기존 텍스트를 덮음
- **원인**: 바운딩 박스가 레이블 영역과 겹침
- **해결**: `validate_form_fields()` / `check_bounding_boxes()`로 겹침 검사 후 수정

### 저장 후 폼 필드가 읽기 전용
- **원인**: 폼 필드 플래그 설정
- **해결**: `writer.update_page_form_field_values()` 후 flatten 옵션 확인

### 구조 추출에서 "(cid:X)" 패턴만 나옴
- **원인**: 스캔/이미지 기반 PDF
- **해결**: Approach B (시각 추정) 사용

### 바운딩 박스 겹침
- **원인**: 레이블과 입력 영역이 겹침
- **해결**: `check_bounding_boxes()` 실행, 레이블/입력 영역 분리

---

## Quick Reference

### pypdf 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `reader.get_fields()` | 모든 폼 필드 조회 |
| `writer.append(reader)` | 전체 PDF 추가 |
| `writer.update_page_form_field_values(page, data)` | 폼 필드 값 업데이트 |
| `page.merge_page(overlay)` | 페이지 병합 (오버레이) |

### reportlab 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `canvas.Canvas(buffer, pagesize)` | 캔버스 생성 |
| `c.setFont(font, size)` | 폰트 설정 |
| `c.setFillColorRGB(r, g, b)` | 색상 설정 |
| `c.drawString(x, y, text)` | 텍스트 그리기 |
| `c.drawCentredString(x, y, text)` | 중앙 정렬 텍스트 |

### 필드 타입 값

| 필드 유형 | 체크/선택 값 | 해제 값 |
|-----------|-------------|---------|
| 체크박스 | `/Yes`, `/On`, `/1` | `/Off`, `/No`, `/0` |
| 라디오 버튼 | `/OptionName` | N/A |
| 드롭다운 | 옵션 텍스트 | N/A |

### fields.json 좌표 체계 선택

| 키 | 좌표 체계 | 사용 시점 |
|----|-----------|----------|
| `pdf_width` / `pdf_height` | PDF 포인트 | 구조 추출 (Approach A) |
| `image_width` / `image_height` | 이미지 픽셀 | 시각 추정 (Approach B) |

### 인라인 함수 목록

| 함수 | 용도 |
|------|------|
| `check_fillable_fields()` | 폼 필드 존재 확인 |
| `get_form_fields()` | Fillable 필드 정보 추출 |
| `fill_pdf_form()` | 단일 페이지 Fillable PDF 채우기 |
| `fill_multipage_form()` | 여러 페이지 Fillable PDF 채우기 |
| `get_checkbox_values()` | 체크박스/라디오 가능한 값 확인 |
| `analyze_form_layout()` | Non-fillable 구조 추출 (pdfplumber) |
| `pdf_to_images()` | PDF -> 이미지 변환 |
| `check_bbox_intersection()` | 두 바운딩 박스 겹침 확인 |
| `validate_form_fields()` | 폼 필드 정의 검증 |
| `check_bounding_boxes()` | fields.json 바운딩 박스 종합 검증 |
| `create_text_overlay()` | reportlab 텍스트 오버레이 생성 |
| `fill_pdf_with_overlay()` | Non-fillable PDF 오버레이 채우기 |
| `pdfplumber_to_reportlab()` | pdfplumber -> reportlab 좌표 변환 |
| `reportlab_to_pdfplumber()` | reportlab -> pdfplumber 좌표 변환 |
| `image_to_pdf_coords()` | 이미지 -> PDF 좌표 변환 |
| `image_to_reportlab_coords()` | 이미지 -> reportlab 좌표 변환 |
| `find_text_and_convert_coords()` | 텍스트 찾기 + 좌표 변환 |
| `register_korean_font()` | 한글 폰트 등록 |
