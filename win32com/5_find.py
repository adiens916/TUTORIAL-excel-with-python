from settings import get_active_sheet_of_excel
from win32com.client import constants

worksheet = get_active_sheet_of_excel('5_find')  # 엑셀 실행

target_cell = worksheet.UsedRange.Find('화곡초')  # '화곡초' 검색
target_cell.Select()  # 선택

target_cell = worksheet.UsedRange.Find('화종초')  # 없는 것 검색
# target_cell.Select()  # 선택하면 에러 남 (None이라서 Select 속성이 없기 때문)
print(target_cell)


