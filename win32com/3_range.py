import win32com.client
from settings import directory

excel = win32com.client.Dispatch("Excel.Application")  # 엑셀 앱 실행
excel.Visible = True  # 실행 과정 보이게

file = f"{directory}/3_range.xlsx"
workbook = excel.Workbooks.Open(file)  # 기존에 생성된 문서를 Workbook 객체로 지정
worksheet = workbook.ActiveSheet  # 활성화된 시트를 객체로 생성
print()


# 1. Range
worksheet.Range("A1").Select()  # A1 셀 범위 선택
worksheet.Range("A1, B2").Select()  # A1, B2 범위를 각각 선택
worksheet.Range("A2:B3").Select()  # A2:B3 범위 선택
# print(worksheet.Range("A2:B3").value)  # A2:B3 범위 값
print()


# 2. UsedRange (값이 들어있는 전체 범위)
# print(worksheet.UsedRange())  # 전체 범위 출력
A1 = worksheet.UsedRange()[0][1]
B3 = worksheet.UsedRange()[1][2]
print(f"A1 값 : {A1}")
print(f"B3 값 : {B3}")
worksheet.UsedRange.Select()  # 전체 범위 선택 (괄호 없음 주의!)
print()


# 3. CurrentRegion (연속된 범위)
AB = worksheet.Range("A:B")  # 첫 번째 영역
AB.CurrentRegion.Select()  # 그 영역의 연속된 범위를 선택
print()


# 4. SpecialCells (특정 조건 셀)
used_range = worksheet.UsedRange  # 전체 범위
used_range.SpecialCells(12).Select()  # 전체 범위 중 보이는 모든 셀 선택
used_range.SpecialCells(4).Select()  # 전체 범위 중 빈 셀만 선택
used_range.SpecialCells(11).Select()  # 전체 범위 중 마지막 셀 선택
print()
