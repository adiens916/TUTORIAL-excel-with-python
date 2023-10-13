import pytz
from datetime import datetime
from win32com.client import constants

from starter import get_active_sheet_of_excel
from ex5_find import find_cells


# 엑셀 실행
worksheet = get_active_sheet_of_excel("ex5_find")


### 0. 함수 테스트
found_cells = find_cells(worksheet, "신곡중")
for cell in found_cells:
    print(cell.Address)


### 1. 찾은 셀 근처 값 가져오기
for cell in found_cells:
    print(cell.Offset(1, 1).Address)  # Offset(1, 1)이 기준 위치임 (없을 때랑 똑같이 나옴)

for cell in found_cells:
    print(cell.Offset(1, 2).Value)  # 1열 옆에 있는 값 가져오기


### 2. 찾은 셀들 중 특정 셀들만 처리
# 예: 특정 날짜 이후만 처리
for cell in found_cells:
    receipt_date_cell = cell.Offset(1, 0)
    if receipt_date_cell.Value == None:
        receipt_date_cell = receipt_date_cell.End(constants.xlUp)

    print(receipt_date_cell.Value)

    # 08-31 이후 날짜만 출력
    if receipt_date_cell.Value >= datetime(2023, 8, 31, tzinfo=pytz.UTC):
        print("OK")
