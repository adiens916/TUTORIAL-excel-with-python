import openpyxl as op

path = "./openpyxl"


def main():
    pass
    # create_new_workbook()
    # load_workbook()


# (A) 새로운 Workbook 객체 생성
def create_new_workbook():
    workbook = op.Workbook()
    # 객체 출력
    print(workbook)

    # 객체 저장
    file_name = "saved.xlsx"
    workbook.save(f"{path}/{file_name}")


# (B) 기존 파일을 객체로 생성하는 경우
def load_workbook():
    file_name = "target.xlsx"
    workbook = op.load_workbook(f"{path}/{file_name}")
    print(workbook)


main()
