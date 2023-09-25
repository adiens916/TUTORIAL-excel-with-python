from settings import get_active_sheet_of_excel


### 1. 셀 한 개 찾기
worksheet = get_active_sheet_of_excel("ex5_find")  # 엑셀 실행

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
        # 다음 셀 찾기
        # 이때 FindNext 안에 찾을 기준이 되는 셀 (이전 셀)을 넣어야 함
        # 안 넣으면 하나만 찾고 끝남 (기준이 없으니 매번 좌측 셀 기준으로만 찾기 때문)
        found_cell = worksheet.UsedRange.FindNext(found_cell)
        print(f"신곡중 위치: {found_cell.Address}")

        # 만족하는 게 없거나 / 전부 다 찾아서 처음으로 돌아온 경우, 종료
        if (found_cell is None) or (found_cell.Address == first_cell_address):
            break


### 2.1. 함수화
def find_cells(keyword: str) -> list:
    found_cells = []

    found_cell = worksheet.UsedRange.Find(keyword)
    if found_cell is not None:
        first_cell_address = found_cell.Address

        while True:
            found_cells.append(found_cell)
            found_cell = worksheet.UsedRange.FindNext(found_cell)

            # 만족하는 게 없거나 / 전부 다 찾아서 처음으로 돌아온 경우, 종료
            if (found_cell is None) or (found_cell.Address == first_cell_address):
                break

    return found_cells
