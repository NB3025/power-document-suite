# Excel 스프레드시트 가이드

Excel 파일 생성, 편집, 수식 사용, 데이터 분석 가이드입니다.

## Overview

Python (openpyxl, pandas)을 사용하여 Excel 파일을 생성, 편집, 분석합니다.

## Workflow Decision Tree

```
Excel 작업 유형?
├── 새 스프레드시트 생성 → openpyxl 사용
├── 기존 파일 편집 → openpyxl (수식 보존)
├── 데이터 분석/변환 → pandas 사용
└── 수식 결과 필요 → data_only=True 또는 LibreOffice
```

## Prerequisites

```bash
pip install openpyxl pandas
```

---

## 출력물 요구사항

### 모든 Excel 파일

- **전문적 폰트**: 일관된 전문 폰트 사용 (Arial, Times New Roman 등)
- **수식 에러 제로**: 모든 Excel 모델은 에러(#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?) 없이 납품
- **기존 템플릿 보존**: 수정 시 기존 포맷, 스타일, 관습을 정확히 매칭

### 재무 모델 색상 규칙

| 색상 | 용도 |
|------|------|
| **파란 텍스트** (0,0,255) | 하드코딩 입력, 시나리오 변경 숫자 |
| **검정 텍스트** (0,0,0) | 모든 수식/계산 |
| **초록 텍스트** (0,128,0) | 같은 워크북 내 다른 시트 링크 |
| **빨간 텍스트** (255,0,0) | 외부 파일 링크 |
| **노란 배경** (255,255,0) | 주요 가정 또는 업데이트 필요 셀 |

### 숫자 포맷 규칙

- **연도**: 텍스트 문자열 ("2024", "2,024" 아님)
- **통화**: $#,##0 형식; 헤더에 단위 명시 ("Revenue ($mm)")
- **0 값**: "-"로 표시 (포맷: `$#,##0;($#,##0);-`)
- **백분율**: 기본 0.0% (소수점 한 자리)
- **배수**: 0.0x (EV/EBITDA, P/E)
- **음수**: 괄호 사용 (123), 마이너스 -123 아님

---

## ⚠️ CRITICAL: 수식 사용 원칙

**Python에서 계산하지 말고, Excel 수식을 사용하세요.**

### ❌ WRONG — Hardcoding

```python
total = df['Sales'].sum()
sheet['B10'] = total  # 5000 하드코딩

growth = (df.iloc[-1]['Revenue'] - df.iloc[0]['Revenue']) / df.iloc[0]['Revenue']
sheet['C5'] = growth  # 0.15 하드코딩
```

### ✅ CORRECT — Excel 수식

```python
sheet['B10'] = '=SUM(B2:B9)'
sheet['C5'] = '=(C4-C2)/C2'
sheet['D20'] = '=AVERAGE(D2:D19)'
```

모든 계산에 적용 — 합계, 백분율, 비율, 차이 등. 소스 데이터 변경 시 자동 재계산 가능해야 함.

---

## ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `sheet['A1'] = sum(list)` | `sheet['A1'] = '=SUM(A2:A10)'` | Python 계산 대신 수식 |
| `load_workbook(f, data_only=True)` 후 저장 | `load_workbook(f)` | data_only 저장시 수식 소실 |
| `sheet['A1'].value = 'ERROR'` | `sheet['A1'] = 'ERROR'` | .value 불필요 |
| 문자열 수식 작은따옴표 | 큰따옴표 또는 따옴표 없음 | `'=SUM(A:A)'` OK |
| 셀 인덱스 0 | 셀 인덱스 1 | openpyxl은 1-indexed |
| 수식 검증 없이 납품 | LibreOffice 재계산 후 에러 확인 | 에러 확인 필수 |

---

## 공통 워크플로우

1. **도구 선택**: pandas (데이터 분석), openpyxl (수식/포맷팅)
2. **생성/로드**: 새 워크북 또는 기존 파일 로드
3. **수정**: 데이터, 수식, 포맷팅 추가/편집
4. **저장**: 파일 쓰기
5. **수식 재계산 (수식 사용 시 필수)**:
   ```bash
   soffice --headless --calc --infilter="Microsoft Excel 2007-2019 XML (.xlsx)" \
     --outdir . --convert-to xlsx output.xlsx
   ```
6. **에러 확인 및 수정**:
   - data_only=True로 열어 에러 셀 확인
   - 에러 수정 후 재계산 반복

---

## 새 스프레드시트 생성

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
sheet = wb.active
sheet.title = "Sales Report"

# 헤더 스타일
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="4472C4")
header_align = Alignment(horizontal="center")

# 헤더 추가
headers = ["제품", "Q1", "Q2", "Q3", "Q4", "합계"]
for col, header in enumerate(headers, 1):
    cell = sheet.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align

# 데이터 추가
data = [
    ["제품 A", 100, 120, 130, 150],
    ["제품 B", 80, 90, 100, 110],
    ["제품 C", 50, 60, 70, 80],
]

for row_idx, row_data in enumerate(data, 2):
    for col_idx, value in enumerate(row_data, 1):
        sheet.cell(row=row_idx, column=col_idx, value=value)
    # 합계 수식 (행별)
    sheet.cell(row=row_idx, column=6, value=f"=SUM(B{row_idx}:E{row_idx})")

# 합계 행
total_row = len(data) + 2
sheet.cell(row=total_row, column=1, value="총계")
for col in range(2, 7):
    col_letter = chr(64 + col)
    sheet.cell(row=total_row, column=col, value=f"=SUM({col_letter}2:{col_letter}{total_row-1})")

# 열 너비
sheet.column_dimensions['A'].width = 15
for col in ['B', 'C', 'D', 'E', 'F']:
    sheet.column_dimensions[col].width = 12

wb.save("sales_report.xlsx")
```

---

## 기존 파일 편집

### 수식 보존하며 편집

```python
from openpyxl import load_workbook

wb = load_workbook("existing.xlsx")  # 수식 보존
sheet = wb.active

sheet['A1'] = "새 값"
sheet.insert_rows(2)
sheet.delete_cols(3)

new_sheet = wb.create_sheet("Summary")
new_sheet['A1'] = "요약 데이터"

wb.save("modified.xlsx")
```

### ⚠️ 주의: data_only=True

```python
# ❌ 위험: 저장하면 수식이 값으로 대체됨
wb = load_workbook("file.xlsx", data_only=True)
# sheet['A1'].value는 계산된 값
# 이 상태로 저장하면 수식 영구 소실!

# ✅ 안전: 수식 보존
wb = load_workbook("file.xlsx")
# sheet['A1'].value는 수식 문자열 (예: "=SUM(B2:B10)")
```

---

## 수식 재계산

openpyxl은 수식을 평가하지 않습니다. 계산된 값이 필요하면:

### 방법 1: LibreOffice 사용

```bash
# 재계산 후 저장
soffice --headless --calc --infilter="Microsoft Excel 2007-2019 XML (.xlsx)" \
  --outdir . --convert-to xlsx file.xlsx
```

### 방법 2: data_only로 읽기 (기존 값)

```python
# 마지막으로 Excel에서 저장된 계산 값 읽기
wb = load_workbook("file.xlsx", data_only=True)
value = sheet['A1'].value  # 계산된 값 (수식 아님)
```

### 방법 3: Python으로 검증

```python
# 수식 결과 확인용 (저장하지 말 것!)
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook("file.xlsx", data_only=True)
ws = wb.active

for row in range(1, 10):
    for col in range(1, 6):
        cell = ws.cell(row=row, column=col)
        print(f"{get_column_letter(col)}{row}: {cell.value}")
```

### 재계산 + 에러 검증 통합

```python
import subprocess
import os
from openpyxl import load_workbook

# 1. LibreOffice로 재계산
def recalc_xlsx(filepath):
    outdir = os.path.dirname(os.path.abspath(filepath))
    subprocess.run([
        "soffice", "--headless", "--calc",
        "--infilter=Microsoft Excel 2007-2019 XML (.xlsx)",
        "--outdir", outdir,
        "--convert-to", "xlsx", filepath
    ], check=True, timeout=60)

# 2. 에러 셀 스캔
def scan_errors(filepath):
    error_types = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!"}
    wb = load_workbook(filepath, data_only=True)
    errors = {}
    total_formulas = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    val = str(cell.value)
                    if val.startswith("="):
                        total_formulas += 1
                    if val in error_types:
                        loc = f"{ws.title}!{cell.coordinate}"
                        errors.setdefault(val, []).append(loc)
    return {
        "total_formulas": total_formulas,
        "total_errors": sum(len(v) for v in errors.values()),
        "error_summary": {k: {"count": len(v), "locations": v} for k, v in errors.items()}
    }

# 사용 예
recalc_xlsx("output.xlsx")
result = scan_errors("output.xlsx")
if result["total_errors"] > 0:
    print("에러 발견:", result["error_summary"])
```

---

## 자주 사용하는 수식

### 기본 수식

```python
# 합계
sheet['F2'] = '=SUM(B2:E2)'

# 평균
sheet['F3'] = '=AVERAGE(B2:E2)'

# 최대/최소
sheet['F4'] = '=MAX(B2:E2)'
sheet['F5'] = '=MIN(B2:E2)'

# 개수
sheet['F6'] = '=COUNT(B2:E2)'
sheet['F7'] = '=COUNTA(B2:E2)'  # 비어있지 않은 셀
```

### 조건부 수식

```python
# IF 조건
sheet['G2'] = '=IF(F2>400,"Good","Bad")'

# 조건부 합계
sheet['G3'] = '=SUMIF(A:A,"제품 A",B:B)'

# 조건부 개수
sheet['G4'] = '=COUNTIF(A:A,"제품 A")'

# 다중 조건 합계
sheet['G5'] = '=SUMIFS(C:C,A:A,"제품 A",B:B,">100")'
```

### 참조 수식

```python
# VLOOKUP
sheet['H2'] = '=VLOOKUP(A2,Data!A:F,2,FALSE)'

# HLOOKUP
sheet['H3'] = '=HLOOKUP("Q1",A1:E10,2,FALSE)'

# INDEX/MATCH (더 유연함)
sheet['H4'] = '=INDEX(B:B,MATCH(A2,A:A,0))'
```

### 백분율/비율

```python
# 백분율
sheet['I2'] = '=B2/SUM(B:B)'

# 전기 대비 성장률
sheet['I3'] = '=(C2-B2)/B2'

# 목표 달성률
sheet['I4'] = '=F2/G2'
```

### 수식 가정 규칙

- 모든 가정(성장률, 마진, 배수 등)은 별도 가정 셀에 배치
- 수식에 하드코딩 값 대신 셀 참조 사용
- 예: `=B5*(1+$B$6)` (O) vs `=B5*1.05` (X)

---

## 수식 검증 체크리스트

### 필수 검증
- [ ] 샘플 2-3개 참조 테스트: 올바른 값을 가져오는지 확인
- [ ] 컬럼 매핑: Excel 컬럼 일치 확인 (예: column 64 = BL)
- [ ] 행 오프셋: Excel은 1-indexed (DataFrame 행 5 = Excel 행 6)

### 일반적 함정
- [ ] NaN 처리: `pd.notna()`로 null 값 확인
- [ ] 0으로 나누기: 분모 확인 (#DIV/0!)
- [ ] 잘못된 참조: 모든 셀 참조가 의도된 셀 가리키는지 확인 (#REF!)
- [ ] 크로스시트 참조: 올바른 형식 사용 (Sheet1!A1)

### 수식 테스트 전략
- [ ] 소규모 시작: 2-3개 셀에서 수식 테스트 후 확장
- [ ] 의존성 확인: 수식에 참조된 모든 셀 존재 확인
- [ ] 엣지 케이스: 0, 음수, 매우 큰 값 포함

---

## 데이터 분석 (pandas)

### 기본 읽기/쓰기

```python
import pandas as pd

# 읽기
df = pd.read_excel('data.xlsx')  # 첫 번째 시트
df = pd.read_excel('data.xlsx', sheet_name='Sheet2')  # 특정 시트
all_sheets = pd.read_excel('data.xlsx', sheet_name=None)  # 모든 시트 (dict)

# 쓰기
df.to_excel('output.xlsx', index=False)

# 여러 시트로 쓰기
with pd.ExcelWriter('multi_sheet.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Data', index=False)
    df2.to_excel(writer, sheet_name='Summary', index=False)
```

### 데이터 분석

```python
# 기본 통계
print(df.describe())
print(df.info())

# 필터링
filtered = df[df['Sales'] > 100]
filtered = df[(df['Sales'] > 100) & (df['Region'] == 'East')]

# 정렬
sorted_df = df.sort_values('Sales', ascending=False)

# 그룹화
grouped = df.groupby('Category')['Sales'].sum()
grouped = df.groupby(['Category', 'Region']).agg({
    'Sales': 'sum',
    'Quantity': 'mean'
})
```

### 피벗 테이블

```python
# 피벗 테이블 생성
pivot = pd.pivot_table(
    df,
    values='Sales',
    index='Category',
    columns='Quarter',
    aggfunc='sum',
    fill_value=0
)

# 마진 추가
pivot_with_totals = pd.pivot_table(
    df,
    values='Sales',
    index='Category',
    columns='Quarter',
    aggfunc='sum',
    margins=True,
    margins_name='Total'
)

# Excel로 저장
pivot.to_excel('pivot_report.xlsx')
```

---

## 스타일링

### 폰트

```python
from openpyxl.styles import Font

# 기본 폰트
cell.font = Font(
    name='Arial',
    size=12,
    bold=True,
    italic=False,
    color='FF0000'  # 빨강 (# 없이!)
)

# 밑줄
cell.font = Font(underline='single')  # single, double
```

### 배경색

```python
from openpyxl.styles import PatternFill

# 단색 배경
cell.fill = PatternFill(
    start_color='FFFF00',  # 노랑
    fill_type='solid'
)

# 그라데이션
cell.fill = PatternFill(
    start_color='FF0000',
    end_color='FFFF00',
    fill_type='linear'
)
```

### 정렬

```python
from openpyxl.styles import Alignment

cell.alignment = Alignment(
    horizontal='center',   # left, center, right
    vertical='center',     # top, center, bottom
    wrap_text=True,        # 텍스트 줄바꿈
    text_rotation=45       # 텍스트 회전
)
```

### 테두리

```python
from openpyxl.styles import Border, Side

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
cell.border = thin_border

# 스타일 옵션: thin, medium, thick, double, dotted, dashed
```

---

## 숫자 포맷

```python
# 통화
cell.number_format = '$#,##0.00'
cell.number_format = '₩#,##0'  # 원화

# 백분율
cell.number_format = '0.0%'
cell.number_format = '0.00%'

# 날짜
cell.number_format = 'YYYY-MM-DD'
cell.number_format = 'YYYY년 MM월 DD일'

# 천 단위 구분
cell.number_format = '#,##0'

# 음수는 괄호
cell.number_format = '#,##0;(#,##0)'

# 음수 괄호 + 0은 대시
cell.number_format = '$#,##0;($#,##0);-'

# 조건부: 양수 초록, 음수 빨강
cell.number_format = '[Green]#,##0;[Red](#,##0)'

# 소수점 자릿수
cell.number_format = '0.000'
```

---

## 차트 생성

```python
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# 막대 차트
chart = BarChart()
chart.title = "Sales by Quarter"
chart.x_axis.title = "Quarter"
chart.y_axis.title = "Sales"

data = Reference(sheet, min_col=2, min_row=1, max_col=5, max_row=4)
categories = Reference(sheet, min_col=1, min_row=2, max_row=4)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)
sheet.add_chart(chart, "H2")

# 선 차트
line_chart = LineChart()
line_chart.title = "Trend"
line_chart.add_data(data, titles_from_data=True)
sheet.add_chart(line_chart, "H15")

# 파이 차트
pie = PieChart()
pie.title = "Market Share"
pie.add_data(Reference(sheet, min_col=2, min_row=1, max_row=4))
pie.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=4))
sheet.add_chart(pie, "H28")
```

---

## 코드 스타일

- 최소한의 간결한 Python 코드 작성
- 불필요한 주석, 장황한 변수명 금지
- 불필요한 print 문 금지
- Excel 파일에는 복잡한 수식이나 가정에 주석 추가
- 하드코딩 값에 데이터 소스 문서화

---

## Troubleshooting

### 에러 타입

| 에러 | 원인 | 해결 |
|------|------|------|
| #REF! | 잘못된 셀 참조 | 셀 주소 확인, 삭제된 행/열 참조 수정 |
| #DIV/0! | 0으로 나눔 | `=IF(B2=0,0,A2/B2)` |
| #VALUE! | 잘못된 값 타입 | 데이터 타입 확인, 텍스트/숫자 혼합 수정 |
| #NAME? | 알 수 없는 함수명 | 함수명 철자 확인, 따옴표 누락 확인 |
| #N/A | 값을 찾을 수 없음 | VLOOKUP 범위 확인, `=IFERROR(VLOOKUP(...),"")` |
| #NULL! | 잘못된 범위 교차 | 범위 연산자 (콜론 vs 공백) 확인 |

### 수식이 계산되지 않음

- **원인**: openpyxl은 수식을 평가하지 않음
- **해결**: LibreOffice로 재계산 또는 Excel에서 열기
  ```bash
  soffice --headless --calc --infilter="Microsoft Excel 2007-2019 XML (.xlsx)" \
    --outdir . --convert-to xlsx output.xlsx
  ```

### 수식이 사라짐

- **원인**: `data_only=True`로 열고 저장함
- **해결**: `data_only=False`로 열기 (기본값)

### 한글 깨짐

- **원인**: 인코딩 문제
- **해결**: UTF-8 사용, `encoding='utf-8-sig'` (BOM 포함)

### 대용량 파일 느림

- **원인**: 전체 파일 메모리 로드
- **해결**: `read_only=True` 또는 `write_only=True` 사용

---

## Quick Reference

### 셀 접근

```python
# 단일 셀
sheet['A1'] = "값"
sheet.cell(row=1, column=1, value="값")

# 범위
for row in sheet['A1:C3']:
    for cell in row:
        print(cell.value)

# 모든 행
for row in sheet.iter_rows(min_row=1, max_row=10):
    print([cell.value for cell in row])
```

### 열 문자 변환

```python
from openpyxl.utils import get_column_letter, column_index_from_string

get_column_letter(1)  # 'A'
get_column_letter(27)  # 'AA'
column_index_from_string('AA')  # 27
```

### 행/열 조작

```python
sheet.insert_rows(2)      # 2번 위치에 행 삽입
sheet.delete_rows(2)      # 2번 행 삭제
sheet.insert_cols(2)      # 2번 위치에 열 삽입
sheet.delete_cols(2)      # 2번 열 삭제
sheet.merge_cells('A1:C1')  # 셀 병합
sheet.unmerge_cells('A1:C1')  # 병합 해제
```

### 시트 조작

```python
wb.create_sheet("NewSheet")  # 새 시트 생성
wb.remove(wb['Sheet1'])      # 시트 삭제
wb.copy_worksheet(sheet)     # 시트 복사
print(wb.sheetnames)         # 시트 목록
```

### 라이브러리 선택

| 도구 | 용도 |
|------|------|
| **pandas** | 데이터 분석, 대량 작업, 간단한 데이터 출력 |
| **openpyxl** | 복잡한 포맷팅, 수식, Excel 고유 기능 |
