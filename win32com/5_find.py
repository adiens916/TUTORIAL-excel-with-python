from settings import get_active_sheet_of_excel
from win32com.client import constants

### 1. 셀 한 개 찾기
worksheet = get_active_sheet_of_excel("5_find")  # 엑셀 실행

target_cell = worksheet.UsedRange.Find("화곡초")  # '화곡초' 검색
target_cell.Select()  # 선택

target_cell = worksheet.UsedRange.Find("화종초")  # 없는 것 검색
# target_cell.Select()  # 선택하면 에러 남 (None이라서 Select 속성이 없기 때문)
print(f"화종초 위치: {target_cell}")


### 2. 셀 여러 개 찾기
found_cell = worksheet.UsedRange.Find("신곡중")
if found_cell is not None:
    first_cell_address = found_cell.Address

    while True:
        found_cell = worksheet.UsedRange.FindNext(found_cell)  # 다음 셀 찾기
        print(f"신곡중 위치: {found_cell.Address}")

        # 만족하는 게 없거나 / 전부 다 찾아서 처음으로 돌아온 경우, 종료
        if (found_cell is None) or (found_cell.Address == first_cell_address):
            break


### 3. 찾은 셀 근처 값 가져오기


### 4. 찾은 셀들 중 특정 셀들만 처리
