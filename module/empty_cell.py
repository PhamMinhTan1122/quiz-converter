import openpyxl

def check(excel_filename, number_question):
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_filename)
    for rowNumber in range(3, number_question):
        if ((wb.active.cell(row=rowNumber, column=7).value)==None):
            print(f"Miss question: {rowNumber - 2}")