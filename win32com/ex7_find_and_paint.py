from starter import get_active_sheet_of_excel
from ex5_find import find_cells
from ex7_color import rgb_to_int


def main():
    # 엑셀 실행
    worksheet = get_active_sheet_of_excel("ex7_color")

    # C열에 'target'이라고 적힌 셀 찾아서 노란 색 칠하기
    # 따로 주소 필요없이, 그냥 Range 타입 참조하면 됨.

    found_cells = find_cells(worksheet, "target")
    for cell in found_cells:
        color = rgb_to_int(255, 255, 0)
        cell.Interior.Color = color


if __name__ == "__main__":
    main()
