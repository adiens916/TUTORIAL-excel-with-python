from starter import get_excel, get_sheet


# 엑셀 실행
workbook = get_excel("ex8_sheet")

# Sheet1 & Sheet2 가져오기
sheet1 = get_sheet(workbook, "Sheet1")
sheet2 = get_sheet(workbook, "Sheet2")

# 각 시트별 데이터 가져오기
A1 = sheet1.Range("A1")
B2 = sheet2.Range("B2")

# 해당 셀을 포함하는 시트명 출력
# 참고: https://learn.microsoft.com/en-us/office/vba/api/excel.range.worksheet
print(f"A1 셀의 시트명: {A1.Worksheet.Name}")
print(f"B2 셀의 시트명: {B2.Worksheet.Name}")
