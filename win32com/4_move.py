import win32com.client
from settings import directory

excel = win32com.client.Dispatch("Excel.Application")  # 엑셀 앱 실행
excel.Visible = True  # 실행 과정 보이게

file = f"{directory}/4_move.xlsx"
workbook = excel.Workbooks.Open(file)  # 기존에 생성된 문서를 Workbook 객체로 지정
worksheet = workbook.ActiveSheet  # 활성화된 시트를 객체로 생성
print()
