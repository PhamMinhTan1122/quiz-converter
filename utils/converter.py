from module.extractor import EXTRACT_DATA
from module.writer import WRITE_TO_EXCEL
import time
class QuizConverter:
    def __init__(
            self, 
            excel_filename, 
            docx_filename, 
            answer_table: bool, 
            is_underline: bool,
            is_italic: bool):
        self.excel_filename = excel_filename
        self.docx_filename = docx_filename
        self.answer_table = answer_table
        self.is_underline = is_underline
        self.is_italic = is_italic
    def convert(self):
        # Print user input information
        print("Excel Filename:", self.excel_filename)
        print("Docx Filename:", self.docx_filename)
        print("Answer Table:", self.answer_table)
        print("Underline:", self.is_underline)
        print("Italic:", self.is_italic)

        # Wait for 3 seconds for user verification
        time.sleep(3)
        # Extract data from the docx file using extractor.py module 
        self.data_list = EXTRACT_DATA(
            docx_filename=self.docx_filename,
            excel_file_name=self.excel_filename, 
            answer_table=self.answer_table, 
            is_underline=self.is_underline,
            is_italic=self.is_italic)


        # Write data to the excel file using writer.py module 
        self.max_number = WRITE_TO_EXCEL(self.data_list, self.excel_filename)
        
