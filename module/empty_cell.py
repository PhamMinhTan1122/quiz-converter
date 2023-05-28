import openpyxl

def check(excel_filename):
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_filename)
    for rowNumber in range(3, 86 + 1):
        if ((wb.active.cell(row=rowNumber, column=7).value)==None):
            print(f"Miss question: {rowNumber - 2}")