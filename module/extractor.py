# This module contains a function to extract data from a docx file
import docx
import re
import openpyxl
def extract_data(docx_filename, excel_file_name, answer_table: bool):
    # Create a document object from the docx file
    doc = docx.Document(docx_filename)
    wb = openpyxl.Workbook()
    wb = openpyxl.load_workbook(excel_file_name)
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
        # print(answer_table)
        table = doc.tables[0]
        # print(table)
        values = []
        for row in table.rows:
            values.append([cell.text for cell in row.cells])
        # Print the values of the table
        for row, items in enumerate(values):
            ws[f"G{int(items[0]) + 2}"] = answer_index[items[1]]
    else:
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
    wb.save(excel_file_name)
    # Return the data_list as the output of this function
    return data_list

def extra_options(data_list, index):
    string = ''
    match = re.match(r"([A-D]\.) (.*)", data_list[index])
    if match:
        string = match.group(2)
    else:
        print("WRONG FORMAT: ", data_list[index])
    return string
def extra_questions(data_list, index):
    string = ''
    match = re.match(r"(^\d+).\s+(.*)$", data_list[index])
    if match:
        # index_question = int(match.group(1))
        string = match.group(2)
    return string