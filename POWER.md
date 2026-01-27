---
name: "power-document-suite"
displayName: "Document Suite"
description: "Office 문서(Word, PowerPoint, Excel, PDF) 생성, 편집, 분석 가이드. Tracked changes, comments, formulas, forms 지원."
keywords: ["docx", "word", "pptx", "powerpoint", "pdf", "xlsx", "excel", "spreadsheet", "tracked changes", "redline", "ooxml", "presentation", "document", "office"]
---

# Document Suite

Office 문서 작업을 위한 Kiro Power입니다.

## Overview

이 Power는 다양한 Office 문서 포맷에 대한 생성, 편집, 분석 가이드를 제공합니다. 각 포맷별로 최적의 라이브러리와 워크플로우를 안내합니다.

## Capabilities Table

| 포맷 | 작업 | 도구 | Steering 파일 |
|------|------|------|---------------|
| Word (.docx) | 새 문서 생성 | docx (JS) | `docx-create.md` |
| Word (.docx) | 기존 문서 편집 | OOXML (Python) | `docx-edit.md` |
| Word (.docx) | Tracked Changes | OOXML (Python) | `docx-edit.md` |
| PowerPoint (.pptx) | 새 프레젠테이션 | html2pptx (JS) | `pptx-create.md` |
| PowerPoint (.pptx) | 템플릿 기반 생성 | OOXML (Python) | `pptx-template.md` |
| PDF | 텍스트/테이블 추출 | pdfplumber (Python) | `pdf-extract.md` |
| PDF | 폼 작성 | pypdf (Python) | `pdf-forms.md` |
| Excel (.xlsx) | 생성/편집/분석 | openpyxl, pandas | `xlsx-guide.md` |

## When to Use This Power

- Word 문서를 프로그래밍 방식으로 생성하거나 편집할 때
- PowerPoint 프레젠테이션을 자동 생성할 때
- PDF에서 데이터를 추출하거나 폼을 작성할 때
- Excel 스프레드시트를 생성하거나 수식을 사용할 때
- Tracked Changes(변경 추적)를 프로그래밍 방식으로 추가할 때

## When NOT to Use This Power

- **단순 텍스트 파일** (.txt, .md) - 일반 파일 작업 사용
- **이미지 편집** - 이미지 처리 도구 사용
- **웹 문서** (HTML/CSS) - 웹 개발 도구 사용
- **데이터베이스 작업** - DB 도구 사용
- **PDF에서 복잡한 레이아웃 보존** - PDF는 추출 목적에 적합, 레이아웃 보존 편집은 제한적

---

# Onboarding

## Step 1: Python 의존성

```bash
# 필수
pip install pypdf pdfplumber reportlab openpyxl pandas defusedxml

# PowerPoint 텍스트 추출용
pip install "markitdown[pptx]"
```

## Step 2: Node.js 의존성

```bash
# Word 문서 생성용
npm install -g docx

# PowerPoint 생성용
npm install -g pptxgenjs playwright sharp
```

## Step 3: 시스템 도구

```bash
# macOS
brew install pandoc poppler libreoffice

# Ubuntu/Linux
sudo apt-get install pandoc poppler-utils libreoffice
```

## Step 4: 의존성 확인

```bash
# 확인 명령어
pandoc --version
pdftotext -v
soffice --version
node -e "require('docx')"
node -e "require('pptxgenjs')"
python -c "import openpyxl; import pdfplumber; print('OK')"
```

---

# Steering 파일 매핑

## Word (.docx)

| 작업 | Steering 파일 | 설명 |
|------|---------------|------|
| 새 문서 생성 | `docx-create.md` | docx.js로 새 Word 문서 생성 |
| 기존 문서 편집 | `docx-edit.md` | OOXML로 기존 문서 수정 |
| Tracked Changes | `docx-edit.md` | 변경 추적 추가 (Redlining) |
| 텍스트 추출 | `docx-edit.md` | pandoc으로 마크다운 변환 |

## PowerPoint (.pptx)

| 작업 | Steering 파일 | 설명 |
|------|---------------|------|
| 새 프레젠테이션 | `pptx-create.md` | HTML→PPTX 워크플로우 |
| 템플릿 기반 생성 | `pptx-template.md` | 기존 템플릿 활용 |
| 텍스트 추출 | `pptx-create.md` | markitdown 사용 |

## PDF

| 작업 | Steering 파일 | 설명 |
|------|---------------|------|
| 텍스트 추출 | `pdf-extract.md` | pdfplumber, pypdf |
| 테이블 추출 | `pdf-extract.md` | pandas로 변환 |
| 병합/분할 | `pdf-extract.md` | pypdf 사용 |
| 폼 작성 | `pdf-forms.md` | 필드 채우기 |

## Excel (.xlsx)

| 작업 | Steering 파일 | 설명 |
|------|---------------|------|
| 스프레드시트 생성 | `xlsx-guide.md` | openpyxl 사용 |
| 수식 사용 | `xlsx-guide.md` | Excel 수식 작성법 |
| 데이터 분석 | `xlsx-guide.md` | pandas 활용 |

---

# Agentic Behavior Guidelines

## 기본 원칙

1. **사용자 요청 분석**: 어떤 포맷의 문서인지, 생성/편집/추출 중 어떤 작업인지 파악
2. **적절한 steering 파일 참조**: 작업에 맞는 steering 파일 읽기
3. **코드 실행 전 확인**: 파일 경로, 의존성 설치 여부 확인
4. **결과 검증**: 생성된 파일이 올바른지 확인

## 문서 생성 시

1. 사용자가 원하는 내용/구조 파악
2. 해당 steering 파일 참조
3. 코드 작성 및 실행
4. 생성된 파일 경로 안내

## 문서 편집 시

1. 원본 파일 경로 확인
2. 어떤 부분을 수정할지 파악
3. 백업 권장 (원본 보존)
4. 편집 후 결과 확인

## 텍스트 추출 시

1. 원본 파일 포맷 확인
2. 추출 방법 선택 (전체/특정 페이지)
3. 추출 결과 표시
4. 필요시 파일로 저장

## "어떻게 하나요?" 질문 시

1. 먼저 어떤 작업인지 명확히 하기
2. 사용 가능한 옵션 설명
3. 권장 방법 제시
4. 사용자 선택 후 실행

---

# Quick Reference

## 텍스트 추출

```bash
# Word → Markdown
pandoc document.docx -o output.md

# Word (Tracked Changes 포함)
pandoc --track-changes=all document.docx -o output.md

# PowerPoint → Text
python -m markitdown presentation.pptx

# PDF → Text
python -c "import pdfplumber; pdf=pdfplumber.open('doc.pdf'); print(pdf.pages[0].extract_text())"
```

## 문서 생성

```bash
# Word (Node.js)
node create-doc.js

# PowerPoint (Node.js)
node create-pptx.js

# PDF (Python)
python create-pdf.py

# Excel (Python)
python create-xlsx.py
```

## 포맷 변환

```bash
# DOCX → PDF
soffice --headless --convert-to pdf document.docx

# PPTX → PDF
soffice --headless --convert-to pdf presentation.pptx

# PDF → Images
pdftoppm -jpeg -r 150 document.pdf page
```

---

# Code Style Guidelines

이 Power를 사용할 때 다음 코드 스타일을 따르세요:

1. **간결한 코드 작성** - 불필요한 주석이나 변수 피하기
2. **print 문 최소화** - 디버깅용만 사용
3. **에러 처리 포함** - 파일 존재 여부 등 확인
4. **경로 절대 경로 사용** - 상대 경로보다 명확

```python
# ✅ GOOD
from pypdf import PdfReader
reader = PdfReader("/path/to/file.pdf")
text = reader.pages[0].extract_text()

# ❌ BAD
from pypdf import PdfReader
# Opening the PDF file for reading
pdf_reader_object = PdfReader("file.pdf")  # 상대 경로
extracted_text_content = pdf_reader_object.pages[0].extract_text()
print("Extracted text:", extracted_text_content)  # 불필요한 출력
```

---

# Troubleshooting

## 의존성 문제

### "Module not found" 오류
```bash
# Python 패키지
pip install <package-name>

# Node.js 패키지
npm install -g <package-name>
```

### "pandoc: command not found"
```bash
# macOS
brew install pandoc

# Ubuntu
sudo apt-get install pandoc
```

### "soffice: command not found"
```bash
# macOS
brew install libreoffice

# Ubuntu
sudo apt-get install libreoffice
```

## 파일 문제

### "Permission denied"
- 파일 권한 확인: `ls -la <file>`
- 쓰기 권한 추가: `chmod +w <file>`

### "File not found"
- 절대 경로 사용 권장
- 파일 존재 확인: `ls <path>`

### 인코딩 문제
- UTF-8 인코딩 사용
- `encoding='utf-8'` 파라미터 추가
