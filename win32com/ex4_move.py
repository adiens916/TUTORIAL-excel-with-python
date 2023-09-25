from settings import get_active_sheet_of_excel
from win32com.client import constants

worksheet = get_active_sheet_of_excel("ex4_move")  # 엑셀 실행

worksheet.Range("A1").Select()  # A1 선택
worksheet.Range("A1").End(constants.xlDown).Select()  # 맨 밑 A9 선택
worksheet.Range("A9").End(constants.xlToRight).Select()  # 맨 우측 D9 선택
print()


# Hint: https://www.mrexcel.com/board/threads/vba-shift-up-arrow-after-ctrl-shift-down-arrow.1082629/
# Ref: https://learn.microsoft.com/en-us/office/vba/api/excel.range.end
