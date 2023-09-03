import openpyxl as op

path = "./openpyxl"


def main():
    pass
    write_cell()


def write_cell():
    workbook = op.load_workbook(f"{path}/target.xlsx")  # 객체 생성
    worksheet = workbook["Sheet1"]  # 객체 생성

    worksheet.cell(1, 1).value = "test 1"  # A1 입력
    worksheet["C1"].value = "test 2"  # C1 입력

    workbook.save(f"{path}/target.xlsx")  # 저장


main()
