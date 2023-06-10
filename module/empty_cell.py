"""
This function checks if there are any unanswered questions in an excel file.

The function imports the following modules:

* `openpyxl`: This module provides a Python interface to Excel files.
* `module.logs`: This module contains a function to write logs to a file.

The `check()` function has the following parameters:

* `excel_filename`: The path to the excel file.
* `number_question`: The number of questions in the excel file.

The `check()` function uses the following steps:

1. Opens the excel file.
2. Iterates through the rows of the excel file.
3. Checks if the cell in column G is empty.
4. If the cell in column G is empty, appends the question number to a list.
5. Writes the list of unanswered questions to a log file.
"""

from openpyxl import load_workbook
from module.logs import Logs

def check(excel_filename, number_question):
    """
    Checks if there are any unanswered questions in an excel file.

    Args:
        excel_filename: The path to the excel file.
        number_question: The number of questions in the excel file.

    Returns:
        A list of unanswered questions.
    """
    # Opens the excel file.
    wb = load_workbook(excel_filename)

    # Opens the excel file.
    logs = []
    for rowNumber in range(3, number_question):
        # Checks if the cell in column G is empty.
        if ((wb.active.cell(row=rowNumber, column=7).value)==None):
            # print(f"Miss question: {rowNumber - 2}")
            logs.append(f"You missed the answer from question number: {rowNumber - 2}")
            
    # Writes the list of unanswered questions to a log file.
    Logs().write_new_line(logs)