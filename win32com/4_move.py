import win32com.client
from win32com.client import constants
from settings import directory

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 엑셀 앱 실행
excel.Visible = True  # 실행 과정 보이게

file = f"{directory}/4_move.xlsx"
workbook = excel.Workbooks.Open(file)  # 기존에 생성된 문서를 Workbook 객체로 지정
worksheet = workbook.ActiveSheet  # 활성화된 시트를 객체로 생성
print()


worksheet.Range("A1").Select()  # A1 선택
worksheet.Range("A1").End(constants.xlDown).Select()  # 맨 밑 A9 선택
worksheet.Range("A9").End(constants.xlToRight).Select()  # 맨 우측 D9 선택
print()


# Hint: https://www.mrexcel.com/board/threads/vba-shift-up-arrow-after-ctrl-shift-down-arrow.1082629/
# Ref: https://learn.microsoft.com/en-us/office/vba/api/excel.range.end