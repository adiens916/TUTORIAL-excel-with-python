import win32com.client

# 엑셀 프로그램 실행
excel = win32com.client.Dispatch("Excel.Application")

# 엑셀 실행과정 보이게
excel.Visible = True
print()

# 엑셀 프로그램에 Workbook 추가 (객체 설정)
# workbook = excel.Workbooks.Add()
# worksheet = workbook.Worksheets("Sheet1")
# print()

# 이미 있는 엑셀 파일 열기
directory = r"C:\Users\adiens\Downloads\Project\TUTORIAL-win32com\win32com"
excel_file_name = r"\1_cell.xlsx"
workbook = excel.Workbooks.Open(directory + excel_file_name)
worksheet = workbook.Worksheets("Sheet1")
print()

# 1. 셀 row, col 값 지정하여 값 넣기
worksheet.Cells(1, 1).Value = "test1"
# 2. Range로 값 넣기
worksheet.Range("A2").Value = "test2"
# 3. Range로 다중범위 지정해서 값 넣기
worksheet.Range("A3:C3").Value = "test3"
# 4. Range로 다중범위 지정하기 다른 버전
worksheet.Range(worksheet.Cells(4, 1), worksheet.Cells(4, 3)).Value = "test4"
print()

# 복사 & 붙여넣기
worksheet.Range("A1:A10").Copy()
worksheet.Range("B1").Select()
worksheet.Paste()
print()
# workbook.Save()
