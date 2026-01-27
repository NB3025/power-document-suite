# Word 문서 편집 (OOXML)

기존 .docx 파일 편집 및 Tracked Changes(변경 추적) 가이드입니다.

## Overview

Word 문서(.docx)는 ZIP으로 압축된 XML 파일들입니다. 기존 문서를 편집하려면 OOXML(Office Open XML)을 직접 수정합니다.

## Workflow Decision Tree

```
Word 문서 작업 유형?
├── 새 문서 생성 → docx-create.md (docx.js)
├── 기존 문서 편집 → 이 가이드 (OOXML)
├── Tracked Changes 추가 → 이 가이드
└── 텍스트 추출 → pandoc --track-changes=all
```

## Prerequisites

```bash
pip install defusedxml
```

---

## ⚠️ CRITICAL: 자주 발생하는 오류

### ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| 변경된 텍스트만 `<w:del>`/`<w:ins>` 안에 | 변경된 부분만 태그 안에, 나머지 밖에 | 최소 범위 마킹 |
| `<w:ins>` 내부 텍스트 직접 수정 | 중첩 `<w:del>` 사용 | 다른 작성자 변경 수정 시 |
| RSID `ABC123` | RSID `00AB1234` (8자리) | 16진수 8자리 |
| `<w:del>` 안에 `<w:t>` | `<w:del>` 안에 `<w:delText>` | 삭제 텍스트 태그 |
| `data_only=True` 저장 | `data_only=False` (기본) | 수식 보존 |
| `<w:ins>...<w:del>` 혼합 | 올바른 태그 닫기 | XML 구조 준수 |

---

## 기본 워크플로우

### Step 1: 언팩 (ZIP 해제)

```bash
unzip document.docx -d unpacked/
```

### Step 2: XML 편집

```
unpacked/
├── [Content_Types].xml        # 콘텐츠 유형 정의
├── word/
│   ├── document.xml           # 본문 내용
│   ├── styles.xml             # 스타일 정의
│   ├── settings.xml           # 문서 설정
│   ├── comments.xml           # 코멘트 (있는 경우)
│   └── _rels/
│       └── document.xml.rels  # 관계 정의
└── _rels/
```

### Step 3: 팩 (ZIP 압축)

```bash
cd unpacked && zip -r ../output.docx . -x "*.DS_Store"
```

---

## 기본 XML 구조

### 단락 (Paragraph)

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r><w:t>텍스트 내용</w:t></w:r>
</w:p>
```

### 텍스트 포맷팅

```xml
<!-- 굵은 텍스트 -->
<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Bold</w:t></w:r>

<!-- 이탤릭 -->
<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>Italic</w:t></w:r>

<!-- 밑줄 -->
<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underlined</w:t></w:r>

<!-- 하이라이트 -->
<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Highlighted</w:t></w:r>
```

### 리스트

```xml
<!-- 번호 리스트 -->
<w:p>
  <w:pPr>
    <w:pStyle w:val="ListParagraph"/>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
    <w:spacing w:before="240"/>
  </w:pPr>
  <w:r><w:t>First item</w:t></w:r>
</w:p>

<!-- 다른 numId = 새 리스트 (1부터 다시 시작) -->
<w:p>
  <w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>
  </w:pPr>
  <w:r><w:t>New list item 1</w:t></w:r>
</w:p>
```

### 테이블

```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="0" w:type="auto"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="4675"/><w:gridCol w:w="4675"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell 1</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell 2</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

---

## Tracked Changes (변경 추적)

### ⚠️ CRITICAL: Tracked Changes 핵심 규칙

1. **변경된 부분만 마킹**: 변경되지 않은 텍스트는 `<w:del>`/`<w:ins>` 외부에 유지
2. **다른 작성자 변경 수정**: 내부 텍스트 직접 수정 금지, 중첩 `<w:del>` 사용
3. **RSID 형식**: 8자리 16진수 (예: `00AB1234`)

### settings.xml에 추적 활성화

```xml
<w:settings>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:trackRevisions/>  <!-- proofState 뒤에 위치 -->
</w:settings>
```

### 텍스트 삽입 (Insertion)

```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-27T00:00:00Z" w16du:dateUtc="2025-01-27T00:00:00Z">
  <w:r w:rsidR="00792858">
    <w:t>추가된 텍스트</w:t>
  </w:r>
</w:ins>
```

### 텍스트 삭제 (Deletion)

```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-27T00:00:00Z" w16du:dateUtc="2025-01-27T00:00:00Z">
  <w:r w:rsidDel="00792858">
    <w:delText>삭제된 텍스트</w:delText>  <!-- ⚠️ delText 사용! -->
  </w:r>
</w:del>
```

### 텍스트 교체 패턴

**"monthly" → "quarterly" 변경 (최소 범위)**:

```xml
<!-- 변경 전: <w:r><w:t>The report is monthly</w:t></w:r> -->

<!-- 변경 후: 변경되지 않은 부분은 외부에 -->
<w:r w:rsidR="00AB12CD"><w:t>The report is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="2025-01-27T00:00:00Z">
  <w:r><w:delText>monthly</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="2025-01-27T00:00:00Z">
  <w:r><w:t>quarterly</w:t></w:r>
</w:ins>
```

### 다른 작성자의 삽입 삭제 (중첩 구조)

```xml
<!-- 다른 사람이 삽입한 "monthly"를 삭제하고 "weekly"로 변경 -->
<!-- ⚠️ CRITICAL: 내부 텍스트 직접 수정 금지! 중첩 del 사용 -->

<w:ins w:author="Jane Smith" w:id="16">
  <w:del w:author="Claude" w:id="40">
    <w:r><w:delText>monthly</w:delText></w:r>
  </w:del>
</w:ins>
<w:ins w:author="Claude" w:id="41">
  <w:r><w:t>weekly</w:t></w:r>
</w:ins>
```

### 다른 작성자의 삭제 복원

```xml
<!-- 다른 사람이 삭제한 내용을 복원 -->
<!-- 원본 del 유지, 새 ins 추가 -->

<w:del w:author="Jane Smith" w:id="50">
  <w:r><w:delText>within 30 days</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="51">
  <w:r><w:t>within 30 days</w:t></w:r>
</w:ins>
```

---

## Comments (코멘트) 추가

### document.xml에 코멘트 범위 지정

```xml
<w:commentRangeStart w:id="0"/>
<w:r><w:t>코멘트 대상 텍스트</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:commentReference w:id="0"/></w:r>
```

### comments.xml 생성/수정

```xml
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Claude" w:date="2025-01-27T00:00:00Z" w:initials="C">
    <w:p>
      <w:r><w:t>코멘트 내용입니다.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
```

### Content_Types.xml 업데이트

```xml
<Override PartName="/word/comments.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
```

### document.xml.rels 업데이트

```xml
<Relationship Id="rId10"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
              Target="comments.xml"/>
```

---

## Python 유틸리티

### 기본 unpack/pack

```python
import zipfile
import os
from defusedxml.minidom import parseString

def unpack_docx(docx_path, output_dir):
    """DOCX 파일을 언팩합니다."""
    with zipfile.ZipFile(docx_path, 'r') as zf:
        zf.extractall(output_dir)

def pack_docx(input_dir, output_path):
    """디렉토리를 DOCX로 팩합니다."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                if file.startswith('.'):  # 숨김 파일 제외
                    continue
                file_path = os.path.join(root, file)
                arc_name = os.path.relpath(file_path, input_dir)
                zf.write(file_path, arc_name)
```

### XML 편집 예제

```python
from defusedxml.minidom import parse
import os

def enable_track_revisions(unpacked_dir):
    """변경 추적을 활성화합니다."""
    settings_path = os.path.join(unpacked_dir, 'word', 'settings.xml')
    dom = parse(settings_path)

    settings = dom.getElementsByTagName('w:settings')[0]

    # trackRevisions 요소 추가
    track_elem = dom.createElement('w:trackRevisions')

    # proofState 뒤에 삽입
    proof_state = dom.getElementsByTagName('w:proofState')
    if proof_state:
        proof_state[0].parentNode.insertBefore(track_elem, proof_state[0].nextSibling)
    else:
        settings.appendChild(track_elem)

    with open(settings_path, 'w', encoding='utf-8') as f:
        dom.writexml(f)

def find_and_replace_tracked(unpacked_dir, old_text, new_text, author="Claude"):
    """텍스트를 찾아 tracked change로 교체합니다."""
    doc_path = os.path.join(unpacked_dir, 'word', 'document.xml')

    with open(doc_path, 'r', encoding='utf-8') as f:
        content = f.read()

    from datetime import datetime
    date = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')

    # 간단한 텍스트 교체 (실제로는 XML 파싱 필요)
    replacement = f'''<w:del w:id="1" w:author="{author}" w:date="{date}">
      <w:r><w:delText>{old_text}</w:delText></w:r>
    </w:del>
    <w:ins w:id="2" w:author="{author}" w:date="{date}">
      <w:r><w:t>{new_text}</w:t></w:r>
    </w:ins>'''

    # 실제 구현에서는 XML 파싱하여 적절한 위치에 삽입
    content = content.replace(f'<w:t>{old_text}</w:t>', replacement)

    with open(doc_path, 'w', encoding='utf-8') as f:
        f.write(content)
```

### 완전한 Tracked Change 예제

```python
from defusedxml.minidom import parse, parseString
import os
import shutil
from datetime import datetime

def add_tracked_change(unpacked_dir, search_text, old_word, new_word, author="Claude"):
    """특정 텍스트 내 단어를 tracked change로 교체합니다."""
    doc_path = os.path.join(unpacked_dir, 'word', 'document.xml')
    dom = parse(doc_path)

    date = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')

    # 모든 w:t 요소 검색
    for t_elem in dom.getElementsByTagName('w:t'):
        text = t_elem.firstChild.nodeValue if t_elem.firstChild else ""

        if search_text in text and old_word in text:
            # 부모 w:r 요소 찾기
            r_elem = t_elem.parentNode
            p_elem = r_elem.parentNode

            # rPr 보존
            rpr_elements = r_elem.getElementsByTagName('w:rPr')
            rpr = rpr_elements[0].toxml() if rpr_elements else ""

            # 텍스트 분할
            parts = text.split(old_word, 1)
            if len(parts) == 2:
                before, after = parts

                # 새 XML 구조 생성
                new_xml = f'''<w:container xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:r>{rpr}<w:t>{before}</w:t></w:r>
                  <w:del w:id="1" w:author="{author}" w:date="{date}">
                    <w:r>{rpr}<w:delText>{old_word}</w:delText></w:r>
                  </w:del>
                  <w:ins w:id="2" w:author="{author}" w:date="{date}">
                    <w:r>{rpr}<w:t>{new_word}</w:t></w:r>
                  </w:ins>
                  <w:r>{rpr}<w:t>{after}</w:t></w:r>
                </w:container>'''

                container = parseString(new_xml).documentElement

                # 원본 r_elem 교체
                for child in list(container.childNodes):
                    p_elem.insertBefore(child.cloneNode(True), r_elem)
                p_elem.removeChild(r_elem)

                break

    with open(doc_path, 'w', encoding='utf-8') as f:
        dom.writexml(f)
```

---

## 하이퍼링크 추가

### 외부 링크

```xml
<!-- document.xml -->
<w:hyperlink r:id="rId5">
  <w:r>
    <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
    <w:t>Link Text</w:t>
  </w:r>
</w:hyperlink>

<!-- word/_rels/document.xml.rels -->
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://www.example.com/"
              TargetMode="External"/>
```

### 내부 링크 (북마크)

```xml
<!-- 링크 -->
<w:hyperlink w:anchor="myBookmark">
  <w:r>
    <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
    <w:t>Go to Section</w:t>
  </w:r>
</w:hyperlink>

<!-- 북마크 대상 -->
<w:bookmarkStart w:id="0" w:name="myBookmark"/>
<w:r><w:t>Target content</w:t></w:r>
<w:bookmarkEnd w:id="0"/>
```

### Hyperlink 스타일 (styles.xml에 필요)

```xml
<w:style w:type="character" w:styleId="Hyperlink">
  <w:name w:val="Hyperlink"/>
  <w:basedOn w:val="DefaultParagraphFont"/>
  <w:rPr>
    <w:color w:val="467886" w:themeColor="hyperlink"/>
    <w:u w:val="single"/>
  </w:rPr>
</w:style>
```

---

## 이미지 삽입

### 파일 복사

```bash
cp image.png unpacked/word/media/image1.png
```

### document.xml에 이미지 추가

```xml
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline>
        <wp:extent cx="2743200" cy="1828800"/>  <!-- EMU 단위 (914400 = 1인치) -->
        <wp:docPr id="1" name="Picture 1"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="image1.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId5"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:ext cx="2743200" cy="1828800"/></a:xfrm>
                <a:prstGeom prst="rect"/>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
```

### document.xml.rels 업데이트

```xml
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

### Content_Types.xml 업데이트

```xml
<Default Extension="png" ContentType="image/png"/>
```

---

## Validation 체크리스트

### XML 검증

- [ ] 모든 태그 올바르게 닫힘
- [ ] `<w:ins>` 는 `</w:ins>` 로 닫힘 (혼합 금지)
- [ ] RSID는 8자리 16진수 (0-9, A-F만)
- [ ] 네임스페이스 선언 있음

### Tracked Changes 검증

- [ ] 변경되지 않은 텍스트는 태그 외부에 있음
- [ ] 삭제 텍스트는 `<w:delText>` 사용 (`<w:t>` 아님)
- [ ] 다른 작성자 변경은 중첩 구조 사용
- [ ] w:id 값이 고유함

### 관계 파일 검증

- [ ] 새 리소스에 대한 Relationship 추가됨
- [ ] Content_Types.xml 업데이트됨
- [ ] rId 참조가 올바름

---

## Troubleshooting

### "Word에서 열 수 없음" 또는 "파일 손상"
- **원인**: XML 구조 오류, 네임스페이스 누락
- **해결**: XML 유효성 검사, 태그 닫힘 확인

### Tracked Changes가 표시되지 않음
- **원인**: settings.xml에 trackRevisions 없음
- **해결**: `<w:trackRevisions/>` 추가

### 삭제된 텍스트가 빨간 줄로 표시 안됨
- **원인**: `<w:t>` 대신 `<w:delText>` 미사용
- **해결**: 삭제 텍스트는 `<w:delText>` 사용

### 코멘트가 표시되지 않음
- **원인**: comments.xml 관계 미등록
- **해결**: document.xml.rels와 Content_Types.xml 업데이트

### 이미지가 표시되지 않음
- **원인**: rId 불일치 또는 미디어 파일 누락
- **해결**: document.xml.rels의 rId 확인, 파일 존재 확인

---

## Quick Reference

### Element Ordering in `<w:pPr>`

```xml
<w:pPr>
  <w:pStyle/>      <!-- 1. 스타일 -->
  <w:numPr/>       <!-- 2. 번호 매기기 -->
  <w:spacing/>     <!-- 3. 간격 -->
  <w:ind/>         <!-- 4. 들여쓰기 -->
  <w:jc/>          <!-- 5. 정렬 -->
</w:pPr>
```

### Character Encoding

| 문자 | Entity |
|------|--------|
| " (여는 쌍따옴표) | `&#8220;` |
| " (닫는 쌍따옴표) | `&#8221;` |
| ' (아포스트로피) | `&#8217;` |
| — (em 대시) | `&#8212;` |

### 공백 보존

```xml
<w:t xml:space="preserve"> text with spaces </w:t>
```
