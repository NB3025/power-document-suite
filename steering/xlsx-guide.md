# Excel 스프레드시트 가이드

Excel 파일 생성, 편집, 수식 사용, 데이터 분석 가이드입니다.

## Overview

이 가이드는 Python (openpyxl, pandas)을 사용하여 Excel 파일을 생성, 편집, 분석하는 방법을 다룹니다.

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

## ⚠️ CRITICAL: 수식 사용 원칙

**Python에서 계산하지 말고, Excel 수식을 사용하세요.**

### ❌ WRONG - Hardcoding Calculated Values

```python
# ❌ Bad: Python에서 계산 후 하드코딩
total = sum(values)
sheet['B10'] = total  # 5000 하드코딩

# ❌ Bad: Python으로 성장률 계산
growth = (df.iloc[-1]['Revenue'] - df.iloc[0]['Revenue']) / df.iloc[0]['Revenue']
sheet['C5'] = growth  # 0.15 하드코딩
```

### ✅ CORRECT - Using Excel Formulas

```python
# ✅ Good: Excel이 계산하도록 수식 사용
sheet['B10'] = '=SUM(B2:B9)'

# ✅ Good: 성장률도 Excel 수식으로
sheet['C5'] = '=(C4-C2)/C2'

# ✅ Good: 평균도 Excel 함수로
sheet['D20'] = '=AVERAGE(D2:D19)'
```

**이유**: 스프레드시트가 소스 데이터 변경 시 자동으로 재계산됩니다.

---

## ❌/✅ Common Mistakes

| ❌ Wrong | ✅ Correct | 설명 |
|----------|-----------|------|
| `sheet['A1'] = sum(list)` | `sheet['A1'] = '=SUM(A2:A10)'` | Python 계산 대신 수식 |
| `load_workbook(f, data_only=True)` 후 저장 | `load_workbook(f)` | data_only 저장시 수식 소실 |
| `sheet['A1'].value = 'ERROR'` | `sheet['A1'] = 'ERROR'` | .value 불필요 |
| 문자열 수식 작은따옴표 | 큰따옴표 또는 따옴표 없음 | `'=SUM(A:A)'` OK |
| 셀 인덱스 0 | 셀 인덱스 1 | openpyxl은 1-indexed |

---

## 새 스프레드시트 생성

### Complete Working Example

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 새 워크북 생성
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

# 합계 행 추가
total_row = len(data) + 2
sheet.cell(row=total_row, column=1, value="총계")
for col in range(2, 7):
    col_letter = chr(64 + col)  # B, C, D, E, F
    sheet.cell(row=total_row, column=col, value=f"=SUM({col_letter}2:{col_letter}{total_row-1})")

# 열 너비 설정
sheet.column_dimensions['A'].width = 15
for col in ['B', 'C', 'D', 'E', 'F']:
    sheet.column_dimensions[col].width = 12

# 저장
wb.save("sales_report.xlsx")
```

---

## 기존 파일 편집

### 수식 보존하며 편집

```python
from openpyxl import load_workbook

# 파일 열기 (수식 보존)
wb = load_workbook("existing.xlsx")
sheet = wb.active

# 셀 값 수정
sheet['A1'] = "새 값"

# 행 삽입 (2번 위치에)
sheet.insert_rows(2)

# 열 삭제 (3번 열)
sheet.delete_cols(3)

# 새 시트 추가
new_sheet = wb.create_sheet("Summary")
new_sheet['A1'] = "요약 데이터"

# 저장
wb.save("modified.xlsx")
```

### ⚠️ 주의: data_only=True

```python
# ❌ 위험: 저장하면 수식이 값으로 대체됨
wb = load_workbook("file.xlsx", data_only=True)
# sheet['A1'].value는 계산된 값
# 이 상태로 저장하면 수식 사라짐!

# ✅ 안전: 수식 보존
wb = load_workbook("file.xlsx")
# sheet['A1'].value는 수식 문자열 (예: "=SUM(B2:B10)")
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

# 조건부: 양수 초록, 음수 빨강
cell.number_format = '[Green]#,##0;[Red](#,##0)'

# 소수점 자릿수
cell.number_format = '0.000'
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
import openpyxl
from openpyxl.utils import get_column_letter

wb = load_workbook("file.xlsx", data_only=True)
ws = wb.active

for row in range(1, 10):
    for col in range(1, 6):
        cell = ws.cell(row=row, column=col)
        print(f"{get_column_letter(col)}{row}: {cell.value}")
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

## Troubleshooting

### 에러 타입

| 에러 | 원인 | 해결 |
|------|------|------|
| #REF! | 잘못된 셀 참조 | 셀 주소 확인 |
| #DIV/0! | 0으로 나눔 | `=IF(B2=0,0,A2/B2)` |
| #VALUE! | 잘못된 값 타입 | 데이터 타입 확인 |
| #NAME? | 알 수 없는 함수명 | 함수명 철자 확인 |
| #N/A | 값을 찾을 수 없음 | VLOOKUP 범위 확인 |

### 수식이 계산되지 않음

- **원인**: openpyxl은 수식을 평가하지 않음
- **해결**: LibreOffice로 재계산 또는 Excel에서 열기

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
