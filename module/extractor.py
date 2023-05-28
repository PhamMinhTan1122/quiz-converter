# This module contains a function to extract data from a docx file
import docx
import re
import openpyxl
def extract_data(docx_filename, excel_file_name):
    # Create a document object from the docx file
    doc = docx.Document(docx_filename)
    wb = openpyxl.load_workbook(excel_file_name)
    ws = wb.active

    # Initialize an empty list to store the extracted data
    data = []

    # Loop through the paragraphs in the document and append their text to the list
    for para in doc.paragraphs:
        text = para.text.strip()
        data.append(text)

        match = re.match(r"(^\d+).\s+(.*)$", text)
        if match:
            row = int(match.group(1)) + 2

        for run in para.runs:
            if run.underline and r"[A-D]\. .*":
                ws[f"G{row}"] = str(run.text).replace(".", "")

    # Join the list into a single string and assign it to the data variable
    data = "\n\n".join(data)

    # Split the data by double newline characters and assign it to the data_list variable
    data_list = data.split("\n\n")

    # Return the data_list as the output of this function
    return data_list

