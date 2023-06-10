"""
This module contains a function to extract data from a docx file and save it to an Excel file.

The `extract_data` function takes three arguments:

* `docx_filename`: The path to the docx file to extract data from.
* `excel_file_name`: The path to the Excel file to save the extracted data to.
* `answer_table`: A boolean value indicating whether the answer key is in a table format (True) or in a paragraph format (False).

The `extract_data` function first creates a `Document` object from the docx file. It then creates a `Workbook` object from the Excel file. It then loops through the paragraphs in the document and appends their text to a list. It then creates a dictionary that maps the answer choices to their corresponding numbers. If the `answer_table` argument is True, the function then extracts the answer key from the table and saves it to the Excel file. Otherwise, the function extracts the answer key from the paragraphs and saves it to the Excel file. The function then joins the list of paragraphs into a single string and splits it by double newline characters. The function then returns the list of paragraphs.

The `extra_options` function takes two arguments:

* `data_list`: The list of paragraphs extracted from the docx file.
* `index`: The index of the paragraph in the list.

The `extra_options` function returns the text of the option at the specified index in the list. If the text of the option does not match the expected format, the function adds a message to a list of errors.

The `check_format_op` function checks the list of errors and prints them to the console if there are any.

The `extra_questions` function takes two arguments:

* `data_list`: The list of paragraphs extracted from the docx file.
* `index`: The index of the paragraph in the list.

The `extra_questions` function returns the text of the question at the specified index in the list.
"""

from openpyxl import load_workbook
from tkinter import messagebox
from openpyxl import Workbook
from module.logs import Logs
from docx import Document
import re


def extract_data(docx_filename, excel_file_name, answer_table: bool):

    # Create a document object from the docx file
    doc = Document(docx_filename)
    wb = Workbook()
    wb = load_workbook(excel_file_name)
    ws = wb.active

    # Initialize an empty list to store the extracted data
    data = []

    # Loop through the paragraphs in the document and append their text to the list
    for para in doc.paragraphs:
        text = para.text.strip()
        data.append(text)

    # List convert A B C D to number
    answer_index = {"A": 1, "B": 2, "C": 3, "D": 4}
    if answer_table:
        # Get the answer key from the table
        table = doc.tables[0]
        values = []
        for row in table.rows:
            values.append([cell.text for cell in row.cells])

        # Save the answer key to the Excel file
        for row, items in enumerate(values):
            ws[f"G{int(items[0]) + 2}"] = answer_index[items[1]]

    else:
        # Loop through the paragraphs in the document and extract the answer key
        for para in doc.paragraphs:
            text = para.text.strip()
            match = re.match(r"(^\d+).\s+(.*)$", text)
            if match:
                row = int(match.group(1)) + 2
            for run in para.runs:
                if run.underline and r"[A-D]\. .*":
                    clean_data_ans = str(run.text).replace(".", "")
                    ws[f"G{row}"] = answer_index[clean_data_ans]
    # Join the list into a single string and assign it to the data variable
    data = "\n\n".join(data)

    # Split the data by double newline characters and assign it to the data_list variable
    data_list = data.split("\n\n")
    try:
        wb.save(excel_file_name)
    except PermissionError:
        messagebox.askokcancel(
            "ERROR", "The excel file is already open. Please close it and try again.")
    # Return the data_list as the output of this function
    return data_list


logs = []


def extra_options(data_list, index):
    """
    Extracts the extra options from the given data list at the given index.

    Args:
        data_list (list): The data list.
        index (int): The index of the extra options.

    Returns:
        str: The extra options.
    """
    string = ''
    match = re.match(r"([A-D]\.) (.*)", data_list[index])
    if match:
        string = match.group(2)
    else:
        logs.append(f"WRONG FORMAT: {data_list[index]}")
    return string


def check_format_op():
    """
    Checks if the extra options are in the wrong format.
    """
    if len(logs) > 0:
        Logs().check_file(logs=logs)
        corr = ["Correct:  A. Something Not A.Something"]
        Logs().write_new_line(logs=corr)


def extra_questions(data_list, index):
    """
    Extracts the extra questions from the given data list at the given index.

    Args:
        data_list (list): The data list.
        index (int): The index of the extra questions.

    Returns:
        str: The extra questions.
    """
    # Check if the index is valid.
    string = ''
    match = re.match(r"(^\d+).\s+(.*)$", data_list[index])
    if match:
        # index_question = int(match.group(1))
        string = match.group(2)
    return string
