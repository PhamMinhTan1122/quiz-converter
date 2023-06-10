import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from utils import converter
from module.logs import Logs
import os

class QuizConverterGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Quiz Converter")
        self.master.iconbitmap(os.path.join(os.path.dirname(__file__), "icons", "icon.ico"))
        # self.master.iconbitmap(os.path.join(os.path.dirname(__file__), "icon.ico"))
        # Create frames
        self.input_frame = tk.Frame(self.master)
        self.input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.button_frame = tk.Frame(self.master)
        self.button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # Create input widgets
        self.docx_label = tk.Label(self.input_frame, text="Select a docx file:")
        self.docx_entry = tk.Entry(self.input_frame)
        self.docx_button = tk.Button(self.input_frame, text="Browse...", command=lambda : self.browse_docx_file(None))
        self.master.bind("<Control-d>", lambda event: self.browse_docx_file(event))

        self.xlsx_label = tk.Label(self.input_frame, text="Select an xlsx file:")
        self.xlsx_entry = tk.Entry(self.input_frame)
        self.xlsx_button = tk.Button(self.input_frame, text="Browse...", command=lambda: self.browse_xlsx_file(None))
        self.master.bind("<Control-e>", lambda event: self.browse_xlsx_file(event))
        self.answer_table_var = tk.BooleanVar()
        self.answer_table_checkbutton = tk.Checkbutton(self.input_frame, text="Answer table format", variable=self.answer_table_var)

        self.docx_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.docx_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.docx_button.grid(row=1, column=2, padx=5, pady=5)

        self.xlsx_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.xlsx_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.xlsx_button.grid(row=3, column=2, padx=5, pady=5)

        self.answer_table_checkbutton.grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=5)

        # Create button widgets
        self.convert_button = tk.Button(self.button_frame, text="Convert", command=lambda: self.convert(None))
        self.master.bind("<Return>", lambda event: self.convert(event))
        self.reset_button = tk.Button(self.button_frame, text="Reset", command=lambda : self.reset_file_path(None))
        self.master.bind("<Control-r>", lambda event: self.reset_file_path(event))
        self.result_label = ScrolledText(self.button_frame, width=60, height=10, wrap=tk.WORD)

        self.convert_button.grid(row=0, column=0, padx=5, pady=5)
        self.reset_button.grid(row=0, column=1, padx=5, pady=5)

        # Configure result label widget
        self.result_label.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.result_label.config(state=tk.DISABLED)

        # Configure grid layout
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(0, weight=1)
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_columnconfigure(1, weight=1)
        self.button_frame.grid_rowconfigure(1, weight=1)




    def browse_docx_file(self, event):
        filename = filedialog.askopenfilename(
            filetypes=[("Document files", "*.docx *.doc")], initialdir=os.getcwd())
        self.docx_entry.delete(0, tk.END)
        self.docx_entry.insert(0, filename)

    def browse_xlsx_file(self, event):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")], initialdir=os.getcwd())
        self.xlsx_entry.delete(0, tk.END)
        self.xlsx_entry.insert(0, filename)

    def convert(self, event):
        docx_filename = self.docx_entry.get()
        xlsx_filename = self.xlsx_entry.get()
        answer_table = self.answer_table_var.get()
        if not docx_filename or not xlsx_filename:
            self.result_label.config(state=tk.NORMAL)
            self.result_label.delete('1.0', tk.END)
            self.result_label.insert(tk.END, "Please select both a docx file and an xlsx file.")
            self.result_label.config(state=tk.DISABLED)
            return
        if not docx_filename.endswith(".docx") or not xlsx_filename.endswith(".xlsx"):
            self.result_label.config(state=tk.NORMAL)
            self.result_label.delete('1.0', tk.END)
            self.result_label.insert(tk.END, "Wrong file format.")
            self.result_label.config(state=tk.DISABLED)
            return
        converter_obj = converter.QuizConverter(docx_filename=docx_filename, excel_filename=xlsx_filename, answer_table=answer_table)
        converter_obj.convert()
        self.result_label.config(state=tk.NORMAL)
        self.result_label.delete('1.0', tk.END)
        self.result_label.insert(tk.END, "Saved completed successfully!\nLogs:\n")
        logs = Logs().read()
        if logs:
            self.result_label.insert(tk.END, logs)
        else:
            self.result_label.insert(tk.END, "No logs found.")
        self.result_label.config(state=tk.DISABLED)

    def reset_file_path(self, event):
        self.docx_entry.delete(0, tk.END)
        self.xlsx_entry.delete(0, tk.END)
        self.result_label.config(state=tk.NORMAL)
        self.result_label.delete('1.0', tk.END)
        self.result_label.insert(tk.END, "No logs found.")
        self.result_label.config(state=tk.DISABLED)