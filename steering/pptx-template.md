# PowerPoint 템플릿 기반 프레젠테이션

기존 템플릿을 활용한 프레젠테이션 생성 및 OOXML 편집 가이드입니다.

## Overview

기존 PowerPoint 파일(.pptx)을 템플릿으로 활용하여 프레젠테이션을 만듭니다. PPTX는 ZIP으로 압축된 XML 파일이며, 직접 편집하여 네이티브 편집 가능한 콘텐츠를 유지합니다.

## Workflow Decision Tree

```
PowerPoint 템플릿 작업?
├── 새 프레젠테이션 생성 (템플릿 없이)
│   └── pptx-create.md (PptxGenJS)
├── 기존 템플릿 기반 생성
│   └── 이 가이드 (OOXML 편집)
└── 기존 PPTX 직접 편집
    └── 이 가이드 (OOXML 편집)
```

## Prerequisites

```bash
# Python 라이브러리
pip install python-pptx "markitdown[pptx]" Pillow defusedxml

# 시스템 도구
brew install poppler  # macOS (pdftoppm)
# sudo apt-get install poppler-utils libreoffice  # Ubuntu
```

---

## 템플릿 기반 워크플로우

### Step 1: 템플릿 분석

```bash
# 텍스트 추출
python -m markitdown template.pptx > template-content.md

# PDF 변환 후 썸네일 생성
soffice --headless --convert-to pdf template.pptx
pdftoppm -jpeg -r 150 template.pdf slide
```

`slide-01.jpg`, `slide-02.jpg` 등으로 레이아웃 확인, markitdown 출력으로 placeholder 텍스트 확인.

### Step 2: 슬라이드 매핑 계획

각 콘텐츠 섹션에 맞는 템플릿 슬라이드를 선택합니다.

**다양한 레이아웃 사용** -- 단조로운 프레젠테이션은 대표적인 실패 패턴. 기본 제목 + bullet 슬라이드만 반복하지 마세요.

적극적으로 활용할 레이아웃:
- 다단 레이아웃 (2단, 3단)
- 이미지 + 텍스트 조합
- 전체 블리드 이미지 + 텍스트 오버레이
- 인용/콜아웃 슬라이드
- 섹션 구분자
- 통계/숫자 콜아웃
- 아이콘 그리드 또는 아이콘 + 텍스트 행

콘텐츠 유형에 맞는 레이아웃 매칭 (예: 핵심 포인트 -> bullet, 팀 정보 -> 다단, 추천사 -> 인용 슬라이드)

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

### Step 3: 언팩

```bash
# PPTX를 디렉토리로 추출
unzip -o template.pptx -d unpacked/
```

스마트 따옴표 이스케이프 (Edit 도구가 ASCII로 변환하기 때문에 사전 처리):

```python
import os, re

def escape_smart_quotes(directory):
    """XML 파일의 스마트 따옴표를 XML 엔티티로 이스케이프합니다."""
    replacements = {
        '\u201c': '&#x201C;',  # "
        '\u201d': '&#x201D;',  # "
        '\u2018': '&#x2018;',  # '
        '\u2019': '&#x2019;',  # '
    }
    for root, dirs, files in os.walk(directory):
        for fname in files:
            if fname.endswith('.xml') or fname.endswith('.rels'):
                fpath = os.path.join(root, fname)
                with open(fpath, 'r', encoding='utf-8') as f:
                    content = f.read()
                original = content
                for char, entity in replacements.items():
                    content = content.replace(char, entity)
                if content != original:
                    with open(fpath, 'w', encoding='utf-8') as f:
                        f.write(content)
                    print(f"  escaped: {fpath}")

escape_smart_quotes('unpacked/')
```

### Step 4: 슬라이드 구조 변경

OOXML 직접 편집으로 슬라이드를 재배열합니다. 더 안정적이고 복제도 지원합니다.

```python
import zipfile
import os
import shutil
from defusedxml import ElementTree as ET

def rearrange_via_ooxml(unpacked_dir, slide_indices):
    """
    OOXML 직접 편집으로 슬라이드를 재배열합니다.
    slide_indices: 0-based 인덱스 리스트 (중복 가능)
    """
    # presentation.xml에서 슬라이드 순서 변경
    pres_path = os.path.join(unpacked_dir, 'ppt', 'presentation.xml')
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
    print(f"Rearranged: {slide_indices}")

# 사용: 슬라이드 0, 2, 2, 5, 24 순서로 (2번 중복)
rearrange_via_ooxml('unpacked/', [0, 2, 2, 5, 24])
```

- 불필요한 슬라이드 삭제 (`<p:sldIdLst>`에서 제거)
- 재사용할 슬라이드 복제 (슬라이드 파일 복사 + 관계 추가)
- `<p:sldIdLst>`에서 슬라이드 순서 변경
- **모든 구조 변경을 Step 5 전에 완료**

슬라이드 복제 시 수동 처리 항목:
1. `ppt/slides/slideN.xml` 복사 (새 번호)
2. `[Content_Types].xml`에 Override 추가
3. `ppt/_rels/presentation.xml.rels`에 Relationship 추가
4. `ppt/presentation.xml`의 `<p:sldIdLst>`에 추가
5. `ppt/slides/_rels/slideN.xml.rels` 복사 (필요시)
6. **노트 슬라이드 참조 제거/수정** (중복 오류 방지)

### Step 5: 콘텐츠 편집

각 `slide{N}.xml`의 텍스트를 업데이트합니다.

**Edit 도구 사용** -- sed나 Python 스크립트 대신 Edit 도구를 사용하세요. 무엇을 어디서 교체하는지 명확하여 더 안정적입니다.

텍스트 인벤토리 추출 (어떤 shape에 어떤 텍스트가 있는지 확인):

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

        # shape를 위치 순서로 정렬 (위->아래, 왼쪽->오른쪽)
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

    print(f"인벤토리 저장: {output_path}")
    return inventory

inventory = extract_text_inventory('template.pptx', 'text-inventory.json')
```

교체 JSON 예시:

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

**교체 JSON 규칙:**
1. **존재하는 shape만 참조**: 인벤토리에 없는 shape ID는 오류 발생
2. **paragraphs 없는 shape는 비워짐**: 명시적으로 내용 유지하려면 paragraphs 포함
3. **bullet 텍스트에 기호 제외**: `- `, `* ` 등 자동 추가됨
4. **bullet: true일 때 level 필수**: 기본값 0
5. **속성은 첫 run에 적용**: 복잡한 포맷팅은 OOXML 직접 편집

텍스트 교체 적용:

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
    print(f"저장 완료: {output_path}")

replace_text('working.pptx', 'replacement-text.json', 'output.pptx')
```

### Step 6: 정리 및 팩

`<p:sldIdLst>`에 없는 슬라이드, 참조되지 않는 미디어, 고아 관계 파일을 정리한 후 다시 ZIP으로 팩합니다.

```python
import os
import re
import shutil

def clean_orphans(unpacked_dir):
    """<p:sldIdLst>에 없는 슬라이드, 미참조 미디어, 고아 관계를 제거합니다."""
    pres_path = os.path.join(unpacked_dir, 'ppt', 'presentation.xml')
    with open(pres_path, 'r', encoding='utf-8') as f:
        pres_xml = f.read()

    # presentation.xml.rels에서 참조된 슬라이드 확인
    rels_path = os.path.join(unpacked_dir, 'ppt', '_rels', 'presentation.xml.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels_xml = f.read()

    # sldIdLst에서 사용 중인 rId 추출
    used_rids = set(re.findall(r'<p:sldId[^>]*r:id="(rId\d+)"', pres_xml))

    # rels에서 슬라이드 관계 추출
    slide_rels = re.findall(
        r'<Relationship Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
        rels_xml
    )

    # 사용되지 않는 슬라이드 파일 제거
    for rid, slide_file in slide_rels:
        if rid not in used_rids:
            slide_path = os.path.join(unpacked_dir, 'ppt', 'slides', slide_file)
            if os.path.exists(slide_path):
                os.remove(slide_path)
                print(f"  removed: {slide_path}")
            # 관련 .rels 파일도 제거
            slide_rels_path = os.path.join(
                unpacked_dir, 'ppt', 'slides', '_rels', slide_file + '.rels'
            )
            if os.path.exists(slide_rels_path):
                os.remove(slide_rels_path)

    print("Orphan cleanup complete.")

clean_orphans('unpacked/')
```

팩:

```bash
# ZIP으로 다시 팩 (unpacked/ -> output.pptx)
cd unpacked/ && zip -r ../output.pptx . -x ".*" && cd ..
```

또는 Python으로:

```python
import zipfile
import os

def pack_pptx(unpacked_dir, output_path):
    """언팩된 디렉토리를 PPTX로 다시 팩합니다."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(unpacked_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arc_name = os.path.relpath(file_path, unpacked_dir)
                zf.write(file_path, arc_name)
    print(f"저장 완료: {output_path}")

pack_pptx('unpacked/', 'output.pptx')
```

### Step 7: 검증 (시각 QA)

```bash
# PDF 변환 후 썸네일 생성
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

`slide-01.jpg`, `slide-02.jpg` 등 생성.

특정 슬라이드만 재렌더링:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

**검증 체크리스트:**
- [ ] 겹치는 요소 없음
- [ ] 텍스트 오버플로우 없음
- [ ] 남은 placeholder 텍스트 없음
- [ ] 슬라이드 가장자리 마진 충분 (>= 0.5")
- [ ] 요소 간 간격 충분 (>= 0.3")
- [ ] 배경과 텍스트 대비 충분
- [ ] 차트/테이블 읽을 수 있는 크기
- [ ] 고아 비주얼 없음 (삭제된 콘텐츠의 이미지/도형)

**Placeholder 텍스트 확인:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

결과가 있으면 수정 후 완료 선언.

---

## 콘텐츠 편집 규칙

### 포맷팅 규칙

- **모든 헤더, 서브헤딩, 인라인 라벨은 볼드**: `<a:rPr>`에 `b="1"` 사용
  - 슬라이드 제목
  - 슬라이드 내 섹션 헤더
  - 인라인 라벨 (예: "Status:", "Description:")
- **유니코드 bullet 사용 금지**: `<a:buChar>` 또는 `<a:buAutoNum>` 사용
- **Bullet 일관성**: 레이아웃에서 상속. `<a:buChar>` 또는 `<a:buNone>`만 지정

---

## Common Pitfalls

### 템플릿 적응

소스 콘텐츠가 템플릿보다 적을 때:
- **남는 요소는 완전히 삭제** (이미지, 도형, 텍스트 박스) -- 텍스트만 비우지 말 것
- 텍스트 삭제 후 고아 비주얼 확인
- 시각 QA로 수량 불일치 확인

다른 길이의 텍스트로 교체할 때:
- **짧은 교체**: 보통 안전
- **긴 교체**: 오버플로우 또는 예기치 않은 줄바꿈 가능
- 텍스트 변경 후 시각 QA
- 템플릿 디자인에 맞게 콘텐츠 잘라내기 또는 분할 고려

**템플릿 슬롯 != 소스 항목**: 템플릿에 4명 팀원이 있고 소스에 3명이면, 4번째의 전체 그룹 삭제 (이미지 + 텍스트 박스), 텍스트만 삭제하지 말 것.

### 멀티 아이템 콘텐츠

여러 항목이면 별도 `<a:p>` 요소 생성 -- **하나의 문자열에 합치지 말 것**.

**WRONG** -- 하나의 단락에 모든 항목:
```xml
<a:p>
  <a:r><a:rPr .../><a:t>Step 1: Do the first thing. Step 2: Do the second thing.</a:t></a:r>
</a:p>
```

**CORRECT** -- 별도 단락, 볼드 헤더:
```xml
<a:p>
  <a:pPr algn="l"><a:lnSpc><a:spcPts val="3919"/></a:lnSpc></a:pPr>
  <a:r><a:rPr lang="en-US" sz="2799" b="1" .../><a:t>Step 1</a:t></a:r>
</a:p>
<a:p>
  <a:pPr algn="l"><a:lnSpc><a:spcPts val="3919"/></a:lnSpc></a:pPr>
  <a:r><a:rPr lang="en-US" sz="2799" .../><a:t>Do the first thing.</a:t></a:r>
</a:p>
<a:p>
  <a:pPr algn="l"><a:lnSpc><a:spcPts val="3919"/></a:lnSpc></a:pPr>
  <a:r><a:rPr lang="en-US" sz="2799" b="1" .../><a:t>Step 2</a:t></a:r>
</a:p>
```

원래 단락의 `<a:pPr>`을 복사하여 줄 간격 보존. 헤더에 `b="1"` 사용.

### 스마트 따옴표

**새 텍스트에 따옴표 추가 시 XML 엔티티 사용:**

```xml
<a:t>the &#x201C;Agreement&#x201D;</a:t>
```

| 문자 | 이름 | 유니코드 | XML 엔티티 |
|------|------|----------|------------|
| " | 왼쪽 큰따옴표 | U+201C | `&#x201C;` |
| " | 오른쪽 큰따옴표 | U+201D | `&#x201D;` |
| ' | 왼쪽 작은따옴표 | U+2018 | `&#x2018;` |
| ' | 오른쪽 작은따옴표 | U+2019 | `&#x2019;` |

### 기타

- **공백 보존**: 앞뒤 공백이 있는 `<a:t>`에 `xml:space="preserve"` 사용
- **XML 파싱**: `defusedxml.minidom` 사용, `xml.etree.ElementTree` 금지 (네임스페이스 손상)

### Common Mistakes

| Wrong | Correct | 설명 |
|-------|---------|------|
| `slides[1]` (첫 슬라이드) | `slides[0]` | 0-indexed |
| bullet 텍스트에 `- ` 포함 | bullet 기호 없이 텍스트만 | 자동 추가됨 |
| `<a:t>` 앞에 `<a:rPr>` 누락 | `<a:r><a:rPr/><a:t>텍스트</a:t></a:r>` | 속성 먼저 |
| 존재하지 않는 shape-99 참조 | 인벤토리에서 실제 shape ID 확인 | 검증 후 교체 |
| 슬라이드 삭제 후 ID 재번호화 | 원래 ID와 파일명 유지 | 참조 유지 |
| `<a:txBody>` 내 순서 오류 | `<a:bodyPr>` -> `<a:lstStyle>` -> `<a:p>` | 순서 필수 |
| 색상에 `#` 포함 | `val="FF0000"` (# 없이) | OOXML 규칙 |
| 유니코드 따옴표 직접 사용 | `&#x201C;` 이스케이프 | ASCII 인코딩 |

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
│   │       ├── slide1.xml.rels
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
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="838200" y="365125"/>
      <a:ext cx="7772400" cy="1470025"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <!-- CRITICAL: 순서 필수 -->
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
<a:r><a:rPr b="1" dirty="0"/><a:t>Bold Text</a:t></a:r>

<!-- 이탤릭 -->
<a:r><a:rPr i="1" dirty="0"/><a:t>Italic Text</a:t></a:r>

<!-- 밑줄 -->
<a:r><a:rPr u="sng" dirty="0"/><a:t>Underlined</a:t></a:r>

<!-- 색상 및 폰트 -->
<a:r>
  <a:rPr sz="2400" dirty="0">
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    <a:latin typeface="Arial"/>
  </a:rPr>
  <a:t>Colored Arial 24pt</a:t>
</a:r>

<!-- 하이라이트 -->
<a:r>
  <a:rPr dirty="0">
    <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
  </a:rPr>
  <a:t>Highlighted</a:t>
</a:r>
```

### 리스트 (Bullet/Numbered)

```xml
<!-- Bullet 리스트 -->
<a:p>
  <a:pPr lvl="0"><a:buChar char="&#8226;"/></a:pPr>
  <a:r><a:rPr dirty="0"/><a:t>첫 번째 항목</a:t></a:r>
</a:p>

<!-- 번호 리스트 -->
<a:p>
  <a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr>
  <a:r><a:rPr dirty="0"/><a:t>첫 번째 항목</a:t></a:r>
</a:p>

<!-- 들여쓰기 (2단계) -->
<a:p>
  <a:pPr lvl="1"><a:buChar char="--"/></a:pPr>
  <a:r><a:rPr dirty="0"/><a:t>하위 항목</a:t></a:r>
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

<!-- 3번 -> 2번 위치로 이동 -->
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
4. **CRITICAL:** `_rels` 파일에서 노트 슬라이드 참조 제거/수정
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

## 단위 변환

| 단위 | 변환 |
|------|------|
| EMU -> 인치 | EMU / 914400 |
| EMU -> 포인트 | EMU / 12700 |
| 인치 -> EMU | 인치 x 914400 |
| 포인트 -> EMU | 포인트 x 12700 |

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

## Troubleshooting

### "PowerPoint에서 열 수 없음" 또는 "파일 손상"
- **원인**: XML 구조 오류, 네임스페이스 누락, 관계 파일 불일치
- **해결**: XML 유효성 검사, `[Content_Types].xml` 및 `_rels` 파일 확인

### 텍스트가 표시되지 않음
- **원인**: `<a:txBody>` 내 요소 순서 오류
- **해결**: `<a:bodyPr>` -> `<a:lstStyle>` -> `<a:p>` 순서 확인

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
- **해결**: UTF-8 인코딩, 유니코드 문자 이스케이프 (" -> `&#x201C;`)

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
