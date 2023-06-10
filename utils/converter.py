"""
This code converts a quiz from a docx file to an excel file.

The code imports the following modules:

* `extractor.py`: This module contains a function to extract data from a docx file.
* `writer.py`: This module contains a function to write data to an excel file.
* `empty_cell.py`: This module contains a function to check if a column in an excel file is empty.

The `QuizConverter` class has the following methods:

* `__init__(self, excel_filename, docx_filename, answer_table)`: This method initializes the class and sets the `excel_filename`, `docx_filename`, and `answer_table` attributes.
* `convert(self)`: This method converts the quiz from the docx file to the excel file.

The `convert()` method uses the following steps:

1. Extracts the data from the docx file.
2. Writes the data to the excel file.
3. Checks if column G is empty and returns the questions in that column.
"""

from module.extractor import extract_data
from module.writer import write_to_excel
from module.empty_cell import check

class QuizConverter:

    def __init__(self, excel_filename, docx_filename, answer_table=False):
        """
        Initializes the class and sets the `excel_filename`, `docx_filename`, and `answer_table` attributes.

        Args:
            excel_filename: The path to the excel file.
            docx_filename: The path to the docx file.
            answer_table: Whether or not the docx file contains an answer table.
        """
        self.excel_filename = excel_filename
        self.docx_filename = docx_filename
        self.answer_table = answer_table
        
    def convert(self):
        """
        Converts the quiz from the docx file to the excel file.

        Returns:
            The number of questions in the excel file.
        """
        # Extract data from the docx file using extractor.py module 
        self.data_list = extract_data(docx_filename=self.docx_filename,excel_file_name=self.excel_filename, answer_table=self.answer_table)
        # Write data to the excel file using writer.py module 
        self.max_number = write_to_excel(self.data_list, self.excel_filename)
        # Check column G is empty and return that questions
        check(self.excel_filename, self.max_number)
        