# This module contains a function to write data to an excel file
import openpyxl
from module.extractor import extra_options
from module.extractor import extra_questions


def write_to_excel(data_list, excel_filename):
    # Create a workbook and a worksheet object
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_filename)
    ws = wb.active

    # Initialize a row counter
    row = 3

    # Loop over the data list with a step of 5
    for i in range(0, len(data_list), 5):
        # Check if the data_list has at least 4 elements from index i
        if len(data_list) >= i + 4:
            # Write the question and its options to the worksheet cells
            ws[f"A{row}"] = extra_questions(data_list, i) # The question
            ws[f"C{row}"] = extra_options(data_list, i + 1)# Option A
            ws[f"D{row}"] = extra_options(data_list, i + 2) # Option B
            ws[f"E{row}"] = extra_options(data_list, i + 3) # Option C
            ws[f"F{row}"] = extra_options(data_list, i + 4) # Option D

            # Increment the row counter by 1
            row += 1

        else:
            print(f"Invalid data format at index {i}")
            # messagebox.showerror(title='ERROR', message=str(f"Invalid data format at index {i}"))

    # Save the workbook as an Excel file with the given filename
    wb.save(excel_filename)
    # print("Saved successful")
    return row
