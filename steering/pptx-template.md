# PowerPoint 템플릿 기반 프레젠테이션

기존 템플릿을 활용한 프레젠테이션 생성 및 OOXML 편집 가이드입니다.

## Overview

PowerPoint 파일(.pptx)은 ZIP으로 압축된 XML 파일들입니다. 기존 템플릿을 활용하여 프레젠테이션을 만들 때는:
1. 템플릿 분석 및 슬라이드 인벤토리 작성
2. 슬라이드 재배열 (복제, 삭제, 순서 변경)
3. 텍스트 교체 및 콘텐츠 수정
4. 검증 및 저장

## Workflow Decision Tree

```
PowerPoint 템플릿 작업?
├── 새 프레젠테이션 생성 (템플릿 없이)
│   └── pptx-create.md (html2pptx)
├── 기존 템플릿 기반 생성
│   ├── Step 1: 템플릿 분석 (썸네일 + 텍스트 추출)
│   ├── Step 2: 인벤토리 작성 (슬라이드 매핑)
│   ├── Step 3: 슬라이드 재배열 (복제/삭제/순서)
│   ├── Step 4: 텍스트 인벤토리 추출
│   ├── Step 5: 교체 JSON 생성
│   ├── Step 6: 텍스트 교체 적용
│   └── Step 7: 검증 (썸네일 확인)
└── 기존 PPTX 직접 편집 (OOXML)
    ├── unzip → XML 편집 → zip
    └── 관계 파일 업데이트 필수
```

## Prerequisites

```bash
# Python 라이브러리
pip install python-pptx defusedxml

# 썸네일 생성용
brew install poppler  # macOS
# sudo apt-get install poppler-utils libreoffice  # Ubuntu
```

---

## ⚠️ CRITICAL: 자주 발생하는 오류

### ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `slides[1]` (첫 슬라이드) | `slides[0]` | 0-indexed |
| bullet 텍스트에 `• ` 포함 | bullet 기호 없이 텍스트만 | 자동 추가됨 |
| `<a:t>` 앞에 `<a:rPr>` 누락 | `<a:r><a:rPr/><a:t>텍스트</a:t></a:r>` | 속성 먼저 |
| 존재하지 않는 shape-99 참조 | 인벤토리에서 실제 shape ID 확인 | 검증 후 교체 |
| 슬라이드 삭제 후 ID 재번호화 | 원래 ID와 파일명 유지 | 참조 유지 |
| `<a:txBody>` 내 순서 오류 | `<a:bodyPr>` → `<a:lstStyle>` → `<a:p>` | 순서 필수 |
| 색상에 `#` 포함 | `val="FF0000"` (# 없이) | OOXML 규칙 |
| 유니코드 따옴표 직접 사용 | `&#8220;` 이스케이프 | ASCII 인코딩 |

---

## 템플릿 기반 워크플로우

### Step 1: 템플릿 분석

```bash
# 텍스트 추출
python -m markitdown template.pptx > template-content.md

# PDF 변환 후 썸네일 생성
soffice --headless --convert-to pdf template.pptx
pdftoppm -jpeg -r 150 template.pdf slide

# 또는 썸네일 그리드 생성 (스크립트 사용 시)
# python scripts/thumbnail.py template.pptx thumbnails --cols 5
```

### Step 2: 템플릿 인벤토리 작성

템플릿의 모든 슬라이드를 분석하여 인벤토리 파일을 작성합니다.

```markdown
# template-inventory.md

**Total Slides: 25**
**IMPORTANT: Slides are 0-indexed (first slide = 0)**

## Title Slides
- Slide 0: Title/Cover - 메인 타이틀, 부제목
- Slide 1: Section Header - 섹션 구분자

## Content Slides
- Slide 2: Title + Body - 제목과 본문
- Slide 3: Two Column - 2단 레이아웃
- Slide 4: Three Column - 3단 레이아웃
- Slide 5: Image + Text - 이미지와 텍스트

## Data Slides
- Slide 10: Chart Placeholder - 차트 영역
- Slide 11: Table Layout - 테이블 레이아웃

## Closing
- Slide 24: Q&A / Contact - 마무리 슬라이드
```

### Step 3: 슬라이드 재배열 (python-pptx)

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from copy import deepcopy
import os

def rearrange_slides(template_path, output_path, slide_indices):
    """
    템플릿에서 특정 슬라이드를 선택하여 새 프레젠테이션을 만듭니다.
    slide_indices: 0-based 인덱스 리스트 (중복 가능)
    """
    prs = Presentation(template_path)

    # 슬라이드 XML 직접 조작 (python-pptx 한계 우회)
    # 실제로는 unzip → 재배열 → zip 방식 권장

    # 간단한 방식: 필요한 슬라이드만 유지
    slides_to_keep = set(slide_indices)
    slides_to_remove = []

    for i, slide in enumerate(prs.slides):
        if i not in slides_to_keep:
            slides_to_remove.append(slide)

    # 슬라이드 삭제 (역순으로)
    for slide in reversed(slides_to_remove):
        rId = prs.part.relate_to(slide.part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[prs.slides.index(slide)]

    prs.save(output_path)
    print(f"✅ 저장 완료: {output_path}")

# 사용: 슬라이드 0, 2, 2, 5, 24 순서로 (2번 중복)
rearrange_slides('template.pptx', 'working.pptx', [0, 2, 2, 5, 24])
```

### OOXML 직접 재배열 (권장)

```python
import zipfile
import os
import shutil
from defusedxml import ElementTree as ET

def rearrange_via_ooxml(template_path, output_path, slide_indices):
    """
    OOXML 직접 편집으로 슬라이드를 재배열합니다.
    더 안정적이고 복제도 지원합니다.
    """
    # 임시 디렉토리에 언팩
    temp_dir = 'temp_pptx'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(template_path, 'r') as zf:
        zf.extractall(temp_dir)

    # presentation.xml에서 슬라이드 순서 변경
    pres_path = os.path.join(temp_dir, 'ppt', 'presentation.xml')
    tree = ET.parse(pres_path)
    root = tree.getroot()

    # 네임스페이스 처리
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    sldIdLst = root.find('.//p:sldIdLst', ns)
    if sldIdLst is not None:
        original_slides = list(sldIdLst)

        # 새 순서로 재배열
        sldIdLst.clear()
        for idx in slide_indices:
            if idx < len(original_slides):
                sldIdLst.append(original_slides[idx])

    tree.write(pres_path, xml_declaration=True, encoding='UTF-8')

    # 다시 팩
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arc_name = os.path.relpath(file_path, temp_dir)
                zf.write(file_path, arc_name)

    shutil.rmtree(temp_dir)
    print(f"✅ 저장 완료: {output_path}")
```

### Step 4: 텍스트 인벤토리 추출

```python
from pptx import Presentation
from pptx.util import Inches
import json

def extract_text_inventory(pptx_path, output_path):
    """모든 텍스트 shape의 상세 정보를 추출합니다."""
    prs = Presentation(pptx_path)
    inventory = {}

    for slide_idx, slide in enumerate(prs.slides):
        slide_key = f"slide-{slide_idx}"
        inventory[slide_key] = {}

        # shape를 위치 순서로 정렬 (위→아래, 왼쪽→오른쪽)
        shapes_with_text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                shapes_with_text.append(shape)

        shapes_with_text.sort(key=lambda s: (s.top, s.left))

        for shape_idx, shape in enumerate(shapes_with_text):
            shape_key = f"shape-{shape_idx}"
            paragraphs = []

            # placeholder 타입 확인
            placeholder_type = None
            if shape.is_placeholder:
                placeholder_type = str(shape.placeholder_format.type).split('.')[-1]

            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs)
                if text.strip():
                    para_info = {"text": text}

                    # 속성 추출 (non-default만)
                    if para.runs:
                        run = para.runs[0]
                        if run.font.bold:
                            para_info["bold"] = True
                        if run.font.italic:
                            para_info["italic"] = True
                        if run.font.size:
                            para_info["font_size"] = run.font.size.pt

                    # bullet 확인
                    if para.level is not None and para.level >= 0:
                        para_info["bullet"] = True
                        para_info["level"] = para.level

                    paragraphs.append(para_info)

            if paragraphs:
                inventory[slide_key][shape_key] = {
                    "placeholder_type": placeholder_type,
                    "left": round(shape.left / Inches(1), 2),
                    "top": round(shape.top / Inches(1), 2),
                    "width": round(shape.width / Inches(1), 2),
                    "height": round(shape.height / Inches(1), 2),
                    "paragraphs": paragraphs
                }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(inventory, f, indent=2, ensure_ascii=False)

    print(f"✅ 인벤토리 저장: {output_path}")
    return inventory

inventory = extract_text_inventory('working.pptx', 'text-inventory.json')
```

### Step 5: 교체 JSON 생성

```json
{
  "slide-0": {
    "shape-0": {
      "paragraphs": [
        {"text": "프로젝트 제안서", "bold": true}
      ]
    },
    "shape-1": {
      "paragraphs": [
        {"text": "2025년 1월"}
      ]
    }
  },
  "slide-1": {
    "shape-0": {
      "paragraphs": [
        {"text": "프로젝트 개요", "bold": true}
      ]
    },
    "shape-1": {
      "paragraphs": [
        {"text": "첫 번째 포인트", "bullet": true, "level": 0},
        {"text": "두 번째 포인트", "bullet": true, "level": 0},
        {"text": "하위 항목", "bullet": true, "level": 1}
      ]
    }
  }
}
```

### ⚠️ CRITICAL: 교체 JSON 규칙

1. **존재하는 shape만 참조**: 인벤토리에 없는 shape ID는 오류 발생
2. **paragraphs 없는 shape는 비워짐**: 명시적으로 내용 유지하려면 paragraphs 포함
3. **bullet 텍스트에 기호 제외**: `• `, `- `, `* ` 등 자동 추가됨
4. **bullet: true일 때 level 필수**: 기본값 0
5. **속성은 첫 run에 적용**: 복잡한 포맷팅은 OOXML 직접 편집

### Step 6: 텍스트 교체 적용

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import json

def replace_text(pptx_path, replacements_path, output_path):
    """JSON 기반으로 텍스트를 교체합니다."""
    prs = Presentation(pptx_path)

    with open(replacements_path, encoding='utf-8') as f:
        replacements = json.load(f)

    for slide_idx, slide in enumerate(prs.slides):
        slide_key = f"slide-{slide_idx}"
        if slide_key not in replacements:
            continue

        # shape 정렬 (인벤토리와 동일하게)
        shapes_with_text = [s for s in slide.shapes if s.has_text_frame]
        shapes_with_text.sort(key=lambda s: (s.top, s.left))

        for shape_idx, shape in enumerate(shapes_with_text):
            shape_key = f"shape-{shape_idx}"
            if shape_key not in replacements[slide_key]:
                # 교체 데이터 없으면 기존 텍스트 비우기
                for para in shape.text_frame.paragraphs:
                    para.clear()
                continue

            new_data = replacements[slide_key][shape_key]
            new_paragraphs = new_data.get("paragraphs", [])

            # 기존 단락 처리
            existing_paras = list(shape.text_frame.paragraphs)

            for i, para_data in enumerate(new_paragraphs):
                if i < len(existing_paras):
                    para = existing_paras[i]
                    para.clear()
                else:
                    para = shape.text_frame.add_paragraph()

                # bullet 설정
                if para_data.get("bullet"):
                    para.level = para_data.get("level", 0)

                # 텍스트 추가
                run = para.add_run()
                run.text = para_data["text"]

                # 포맷팅 적용
                if para_data.get("bold"):
                    run.font.bold = True
                if para_data.get("italic"):
                    run.font.italic = True
                if para_data.get("font_size"):
                    run.font.size = Pt(para_data["font_size"])
                if para_data.get("color"):
                    # RRGGBB 형식
                    color = para_data["color"]
                    run.font.color.rgb = RGBColor(
                        int(color[0:2], 16),
                        int(color[2:4], 16),
                        int(color[4:6], 16)
                    )

            # 남은 기존 단락 비우기
            for j in range(len(new_paragraphs), len(existing_paras)):
                existing_paras[j].clear()

    prs.save(output_path)
    print(f"✅ 저장 완료: {output_path}")

replace_text('working.pptx', 'replacement-text.json', 'output.pptx')
```

### Step 7: 검증 (썸네일 확인)

```bash
# PDF 변환 후 썸네일 생성
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf final-slide

# 검증 체크리스트:
# - [ ] 텍스트가 잘리지 않음
# - [ ] 텍스트가 겹치지 않음
# - [ ] 레이아웃이 유지됨
# - [ ] 폰트/색상이 올바름
```

---

## OOXML 구조

### PowerPoint 파일 구조

```
unpacked/
├── [Content_Types].xml        # 콘텐츠 유형 정의
├── _rels/
│   └── .rels                  # 최상위 관계
├── ppt/
│   ├── presentation.xml       # 프레젠테이션 메타데이터
│   ├── _rels/
│   │   └── presentation.xml.rels  # 슬라이드 관계
│   ├── slides/
│   │   ├── slide1.xml         # 슬라이드 내용
│   │   ├── slide2.xml
│   │   └── _rels/
│   │       ├── slide1.xml.rels   # 슬라이드별 관계
│   │       └── slide2.xml.rels
│   ├── slideLayouts/          # 레이아웃 템플릿
│   ├── slideMasters/          # 마스터 슬라이드
│   ├── theme/                 # 테마 정의
│   └── media/                 # 이미지, 미디어
└── docProps/
    ├── app.xml                # 앱 속성
    └── core.xml               # 코어 메타데이터
```

### 슬라이드 XML 구조

```xml
<!-- ppt/slides/slide1.xml -->
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>...</p:nvGrpSpPr>
      <p:grpSpPr>...</p:grpSpPr>
      <!-- Shapes -->
      <p:sp>...</p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

### 텍스트 Shape 구조

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1"/>
    </p:cNvSpPr>
    <p:nvPr>
      <p:ph type="title"/>  <!-- placeholder 타입 -->
    </p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="838200" y="365125"/>  <!-- EMU 단위 -->
      <a:ext cx="7772400" cy="1470025"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <!-- ⚠️ CRITICAL: 순서 필수 -->
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" dirty="0"/>
        <a:t>텍스트 내용</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" dirty="0"/>
    </a:p>
  </p:txBody>
</p:sp>
```

### 텍스트 포맷팅

```xml
<!-- 굵은 텍스트 -->
<a:r>
  <a:rPr b="1" dirty="0"/>
  <a:t>Bold Text</a:t>
</a:r>

<!-- 이탤릭 -->
<a:r>
  <a:rPr i="1" dirty="0"/>
  <a:t>Italic Text</a:t>
</a:r>

<!-- 밑줄 -->
<a:r>
  <a:rPr u="sng" dirty="0"/>
  <a:t>Underlined</a:t>
</a:r>

<!-- 색상 및 폰트 -->
<a:r>
  <a:rPr sz="2400" dirty="0">
    <a:solidFill>
      <a:srgbClr val="FF0000"/>  <!-- # 없이! -->
    </a:solidFill>
    <a:latin typeface="Arial"/>
  </a:rPr>
  <a:t>Colored Arial 24pt</a:t>
</a:r>

<!-- 하이라이트 -->
<a:r>
  <a:rPr dirty="0">
    <a:highlight>
      <a:srgbClr val="FFFF00"/>
    </a:highlight>
  </a:rPr>
  <a:t>Highlighted</a:t>
</a:r>
```

### 리스트 (Bullet/Numbered)

```xml
<!-- Bullet 리스트 -->
<a:p>
  <a:pPr lvl="0">
    <a:buChar char="•"/>
  </a:pPr>
  <a:r>
    <a:rPr dirty="0"/>
    <a:t>첫 번째 항목</a:t>
  </a:r>
</a:p>

<!-- 번호 리스트 -->
<a:p>
  <a:pPr lvl="0">
    <a:buAutoNum type="arabicPeriod"/>
  </a:pPr>
  <a:r>
    <a:rPr dirty="0"/>
    <a:t>첫 번째 항목</a:t>
  </a:r>
</a:p>

<!-- 들여쓰기 (2단계) -->
<a:p>
  <a:pPr lvl="1">
    <a:buChar char="–"/>
  </a:pPr>
  <a:r>
    <a:rPr dirty="0"/>
    <a:t>하위 항목</a:t>
  </a:r>
</a:p>
```

---

## 슬라이드 작업 (OOXML)

### 슬라이드 추가

1. `ppt/slides/slideN.xml` 생성
2. `[Content_Types].xml`에 Override 추가
3. `ppt/_rels/presentation.xml.rels`에 Relationship 추가
4. `ppt/presentation.xml`의 `<p:sldIdLst>`에 추가
5. `ppt/slides/_rels/slideN.xml.rels` 생성 (필요시)
6. `docProps/app.xml` 슬라이드 수 업데이트

### 슬라이드 순서 변경

```xml
<!-- ppt/presentation.xml -->
<!-- 원본 순서 -->
<p:sldIdLst>
  <p:sldId id="256" r:id="rId2"/>
  <p:sldId id="257" r:id="rId3"/>
  <p:sldId id="258" r:id="rId4"/>
</p:sldIdLst>

<!-- 3번 → 2번 위치로 이동 -->
<p:sldIdLst>
  <p:sldId id="256" r:id="rId2"/>
  <p:sldId id="258" r:id="rId4"/>  <!-- 이동됨 -->
  <p:sldId id="257" r:id="rId3"/>
</p:sldIdLst>
```

### 슬라이드 삭제

1. `ppt/presentation.xml`에서 `<p:sldId>` 제거
2. `ppt/_rels/presentation.xml.rels`에서 Relationship 제거
3. `[Content_Types].xml`에서 Override 제거
4. `ppt/slides/slideN.xml` 및 `_rels` 파일 삭제
5. `docProps/app.xml` 업데이트
6. 미사용 미디어 정리

**주의:** 슬라이드 번호를 재부여하지 마세요. 원래 ID와 파일명을 유지하세요.

### 슬라이드 복제

1. 소스 슬라이드 XML 복사
2. 모든 ID를 고유 값으로 변경
3. "슬라이드 추가" 절차 수행
4. **⚠️ CRITICAL:** `_rels` 파일에서 노트 슬라이드 참조 제거/수정
5. 미사용 미디어 참조 제거

---

## 파일 업데이트 체크리스트

### [Content_Types].xml

```xml
<!-- 슬라이드 추가 시 -->
<Override PartName="/ppt/slides/slide3.xml"
          ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>

<!-- 이미지 추가 시 -->
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpg" ContentType="image/jpeg"/>
```

### ppt/_rels/presentation.xml.rels

```xml
<!-- 슬라이드 관계 -->
<Relationship Id="rId2"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
              Target="slides/slide1.xml"/>

<!-- 마스터 슬라이드 관계 -->
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
              Target="slideMasters/slideMaster1.xml"/>
```

### ppt/slides/_rels/slideN.xml.rels

```xml
<!-- 레이아웃 참조 -->
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
              Target="../slideLayouts/slideLayout1.xml"/>

<!-- 이미지 참조 -->
<Relationship Id="rId2"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="../media/image1.png"/>
```

### docProps/app.xml

```xml
<Slides>5</Slides>
<Paragraphs>25</Paragraphs>
<Words>150</Words>
```

---

## Placeholder 타입

| 타입 | 설명 | 용도 |
|------|------|------|
| `ctrTitle` | 중앙 제목 | 타이틀 슬라이드 메인 제목 |
| `subTitle` | 부제목 | 타이틀 슬라이드 부제 |
| `title` | 제목 | 일반 슬라이드 제목 |
| `body` | 본문 | 콘텐츠 영역 |
| `dt` | 날짜 | 날짜 필드 |
| `ftr` | 푸터 | 푸터 텍스트 |
| `sldNum` | 슬라이드 번호 | 페이지 번호 |

---

## 단위 변환

| 단위 | 변환 |
|------|------|
| EMU → 인치 | EMU / 914400 |
| EMU → 포인트 | EMU / 12700 |
| 인치 → EMU | 인치 × 914400 |
| 포인트 → EMU | 포인트 × 12700 |

```python
def emu_to_inches(emu):
    return emu / 914400

def inches_to_emu(inches):
    return int(inches * 914400)

def emu_to_pt(emu):
    return emu / 12700

def pt_to_emu(pt):
    return int(pt * 12700)
```

---

## Troubleshooting

### "PowerPoint에서 열 수 없음" 또는 "파일 손상"
- **원인**: XML 구조 오류, 네임스페이스 누락, 관계 파일 불일치
- **해결**: XML 유효성 검사, `[Content_Types].xml` 및 `_rels` 파일 확인

### 텍스트가 표시되지 않음
- **원인**: `<a:txBody>` 내 요소 순서 오류
- **해결**: `<a:bodyPr>` → `<a:lstStyle>` → `<a:p>` 순서 확인

### 슬라이드 순서가 맞지 않음
- **원인**: `presentation.xml`의 `<p:sldIdLst>` 순서
- **해결**: `<p:sldId>` 요소 순서 조정

### 이미지가 표시되지 않음
- **원인**: `_rels` 파일에 관계 누락 또는 미디어 파일 없음
- **해결**: Relationship 추가 및 `ppt/media/` 파일 확인

### 스타일이 적용되지 않음
- **원인**: `<a:rPr>`가 `<a:t>` 뒤에 있거나 누락
- **해결**: `<a:r><a:rPr .../><a:t>...</a:t></a:r>` 순서 확인

### 템플릿 복제 후 오류
- **원인**: 노트 슬라이드 참조 중복, 미사용 미디어 참조
- **해결**: `_rels` 파일에서 중복/누락 참조 제거

### 한글/특수문자 깨짐
- **원인**: 인코딩 문제 또는 이스케이프 누락
- **해결**: UTF-8 인코딩, 유니코드 문자 이스케이프 (`"` → `&#8220;`)

---

## 검증 체크리스트

### 팩 전 확인

- [ ] 모든 슬라이드가 `[Content_Types].xml`에 선언됨
- [ ] 모든 슬라이드가 `presentation.xml.rels`에 관계 있음
- [ ] `<p:sldIdLst>`에 모든 슬라이드 ID 포함
- [ ] `_rels` 파일에 삭제된 리소스 참조 없음
- [ ] 미사용 미디어 파일 정리됨
- [ ] 노트 슬라이드 참조 정리됨

### XML 검증

- [ ] 모든 태그 올바르게 닫힘
- [ ] 네임스페이스 선언 있음
- [ ] `<a:txBody>` 내 순서 올바름
- [ ] 색상 값에 `#` 없음
- [ ] 유니코드 문자 이스케이프됨

---

## Quick Reference

### python-pptx 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `Presentation(path)` | PPTX 열기 |
| `prs.slides` | 슬라이드 컬렉션 |
| `slide.shapes` | shape 컬렉션 |
| `shape.has_text_frame` | 텍스트 포함 여부 |
| `shape.text_frame.paragraphs` | 단락 리스트 |
| `para.add_run()` | 텍스트 run 추가 |
| `prs.save(path)` | 저장 |

### OOXML 주요 요소

| 요소 | 설명 |
|------|------|
| `<p:sld>` | 슬라이드 루트 |
| `<p:spTree>` | shape 트리 |
| `<p:sp>` | shape |
| `<p:txBody>` | 텍스트 본문 |
| `<a:p>` | 단락 |
| `<a:r>` | 텍스트 run |
| `<a:rPr>` | run 속성 |
| `<a:t>` | 텍스트 내용 |

### 텍스트 속성

| 속성 | 값 | 설명 |
|------|-----|------|
| `b` | `"1"` | 굵게 |
| `i` | `"1"` | 기울임 |
| `u` | `"sng"` | 밑줄 |
| `sz` | `"2400"` | 폰트 크기 (1/100 pt) |
| `dirty` | `"0"` | 정리 상태 |
