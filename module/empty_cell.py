import openpyxl
from tkinter import messagebox

def check(excel_filename, number_question):
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_filename)
    error_shown = False
    for rowNumber in range(3, number_question):
        if ((wb.active.cell(row=rowNumber, column=7).value)==None):
            if not error_shown:
                print(f"Miss question: {rowNumber - 2}")
                messagebox.showerror(title='ERROR', message=f"Miss question: {rowNumber - 2}")
                error_shown = True