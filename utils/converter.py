from module.extractor import extract_data
from module.writer import write_to_excel
from module.empty_cell import check
class QuizConverter:
    def __init__(self, excel_filename, docx_filename):
        self.excel_filename = excel_filename
        self.docx_filename = docx_filename

    def convert(self):
        # Extract data from the docx file using extractor.py module 
        self.data_list = extract_data(docx_filename=self.docx_filename,excel_file_name=self.excel_filename)

        # Write data to the excel file using writer.py module 
        self.max_number = write_to_excel(self.data_list, self.excel_filename)
        # Check column G is empty and return that questions
        check(self.excel_filename, self.max_number)
        