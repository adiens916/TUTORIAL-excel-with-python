import os
import win32com.client

root = os.getcwd()
directory = root + "/win32com"


def get_active_sheet_of_excel(name: str):
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 엑셀 앱 실행
    excel.Visible = True  # 실행 과정 보이게

    file = f"{directory}/{name}.xlsx"
    workbook = excel.Workbooks.Open(file)  # 기존에 생성된 문서를 Workbook 객체로 지정
    worksheet = workbook.ActiveSheet  # 활성화된 시트를 객체로 생성
    return worksheet


def get_excel(excel_name: str):
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 엑셀 앱 실행
    excel.Visible = True  # 실행 과정 보이게

    file = f"{directory}/{excel_name}.xlsx"
    workbook = excel.Workbooks.Open(file)  # 기존에 생성된 문서를 Workbook 객체로 지정
    return workbook


def get_sheet(workbook, sheet_name: str):
    worksheet = workbook.Worksheets(sheet_name)  # 특정 시트명으로 접근
    return worksheet
