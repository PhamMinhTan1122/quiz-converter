
import tkinter as tk
from tkinter import filedialog
from utils import converter

class QuizConverterGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Quiz Converter")

        # Create widgets
        self.docx_label = tk.Label(self.master, text="Select a docx file:")
        self.docx_entry = tk.Entry(self.master)
        self.docx_button = tk.Button(self.master, text="Browse...", command=self.browse_docx_file)

        self.xlsx_label = tk.Label(self.master, text="Select an xlsx file:")
        self.xlsx_entry = tk.Entry(self.master)
        self.xlsx_button = tk.Button(self.master, text="Browse...", command=self.browse_xlsx_file)

        self.answer_table_var = tk.BooleanVar()
        self.answer_table_checkbutton = tk.Checkbutton(self.master, text="Answer table format", variable=self.answer_table_var)

        self.convert_button = tk.Button(self.master, text="Convert", command=self.convert)
        self.result_label = tk.Label(self.master, text="")
        

        # Layout widgets
        self.docx_label.grid(row=0, column=0, sticky="w")
        self.docx_entry.grid(row=0, column=1, padx=10, pady=10)
        self.docx_button.grid(row=0, column=2)

        self.xlsx_label.grid(row=1, column=0, sticky="w")
        self.xlsx_entry.grid(row=1, column=1, padx=10, pady=10)
        self.xlsx_button.grid(row=1, column=2)

        self.answer_table_checkbutton.grid(row=2, column=1, sticky="w")

        self.convert_button.grid(row=3, column=1, pady=10)
        self.result_label.grid(row=4, column=0, columnspan=3)
    def ouput_exc(self, text):
        self.result_label.config(text=text)

    def browse_docx_file(self):
        filename = filedialog.askopenfilename()
        self.docx_entry.delete(0, tk.END)
        self.docx_entry.insert(0, filename)

    def browse_xlsx_file(self):
        filename = filedialog.askopenfilename()
        self.xlsx_entry.delete(0, tk.END)
        self.xlsx_entry.insert(0, filename)

    def convert(self):
        docx_filename = self.docx_entry.get()
        xlsx_filename = self.xlsx_entry.get()
        answer_table = self.answer_table_var.get()
        if not docx_filename or not xlsx_filename:
            self.result_label.config(text="Please select both a docx file and an xlsx file.")
            return
        if not docx_filename.endswith(".docx") or not xlsx_filename.endswith(".xlsx"):
            self.result_label.config(text="Wrong file format.")
            return
        converter_obj = converter.QuizConverter(docx_filename=docx_filename, excel_filename=xlsx_filename, answer_table=answer_table)
        converter_obj.convert()
        self.result_label.config(text="Conversion completed successfully!")