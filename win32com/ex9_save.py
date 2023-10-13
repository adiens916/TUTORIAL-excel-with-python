from random import random
from settings import get_excel, get_sheet


# 엑셀 실행
workbook = get_excel("ex9_save")
worksheet = get_sheet(workbook, "Sheet1")

# A1 셀에 임의 값 입력
worksheet.Range("A1").Value = random()

# 저장
workbook.Save()
