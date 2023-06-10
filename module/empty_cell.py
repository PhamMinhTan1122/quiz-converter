import openpyxl
from module.logs import Logs
def check(excel_filename, number_question):
    logs = []
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_filename)
    for rowNumber in range(3, number_question):
        if ((wb.active.cell(row=rowNumber, column=7).value)==None):
            print(f"Miss question: {rowNumber - 2}")
            logs.append(f"Miss question: {rowNumber - 2}")
    Logs().check_file(logs)