from starter import get_active_sheet_of_excel


# 엑셀 실행
worksheet = get_active_sheet_of_excel("ex7_color")

# 1. 셀 색 값 출력
# A열 기본 색 값 출력
for i in range(1, 7):
    cell = worksheet.Range(f"A{i}")
    color_index = cell.Interior.ColorIndex
    color = cell.Interior.Color
    print(f"index: {color_index} = rgb: {color}")


# 2. 셀 색 RGB 값으로 넣기
# 참고: https://stackoverflow.com/questions/11444207/setting-a-cells-fill-rgb-color-with-pywin32-in-excel
def rgb_to_int(red: int, green: int, blue: int) -> int:
    return red + green * 256 + blue * 256 * 256


# B열에 색 넣기
for i in range(1, 6):
    color = rgb_to_int(50 * i, 50 * i, 50 * i)
    worksheet.Range(f"B{i}").Interior.Color = color
