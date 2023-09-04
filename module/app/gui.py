from tkinter import Frame, Label, Entry, Button, BooleanVar, Checkbutton
from tkinter import DISABLED, END, WORD, NORMAL
from tkinter.scrolledtext import ScrolledText
from webbrowser import open_new_tab 
from tkinter import filedialog
from module.logs import Logs
from utils import converter
import os


class QuizConverterGUI:
    """This class represents a GUI for converting quiz files from docx to xlsx format.

    Args:
        master: The parent window.
    """
    
    def __init__(self, master):
        """Initializes the GUI.

        Args:
            master: The parent window.
        """ 
        self.master = master

        # Set the window title.
        self.master.title("Quiz Converter by Pham Minh Tan")

        # Set the window icon.
        self.master.iconbitmap(os.path.join(os.path.dirname(__file__), "icons", "icon.ico"))
        # print(os.path.join(os.path.dirname(__file__), "icons", "icon.ico"))
        # Create frames
        self.input_frame = Frame(self.master)
        self.input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.button_frame = Frame(self.master)
        self.button_frame.grid(row=1, column=0, padx=10,
                               pady=10, sticky="nsew")
        # Link profile
        self.link_you = Label(self.master, text="Youtube",
                              fg="blue", cursor="hand2")
        self.link_you.grid(row=5, column=0, sticky="nsew")

        # Bind a click event to the link label to open the Youtube channel in a new browser tab.
        self.link_you.bind("<Button-1>", lambda e: self.open_browser(
            url="https://www.youtube.com/channel/UCXM_mZ8NYxuGvDAikcNei8w"))
        self.link_red = Label(self.master, text="Reddit",
                              fg="blue", cursor="hand2")
        self.link_red.grid(row=6, column=0, sticky="nsew")

        # Bind a click event to the link label to open the Reddit profile in a new browser tab.
        self.link_red.bind("<Button-1>", lambda e: self.open_browser(
            url="https://www.reddit.com/user/master_minh_tan"))
        self.copyright = Label(
            self.master, text="Powered by Pham Minh Tan", font=("TkDefaultFont", 7))
        self.copyright.grid(row=7, column=0, sticky="nsew")

        # Create input widgets
        self.docx_label = Label(self.input_frame, text="Select a docx file:")
        self.docx_entry = Entry(self.input_frame)
        self.docx_button = Button(
            self.input_frame, text="Browse...", command=lambda: self.browse_docx_file(None))
        
        # Bind a keyboard shortcut to the browse button.
        self.master.bind("<Control-d>", lambda event: self.browse_docx_file(event))

        self.xlsx_label = Label(self.input_frame, text="Select an xlsx file:")
        self.xlsx_entry = Entry(self.input_frame)
        self.xlsx_button = Button(
            self.input_frame, text="Browse...", command=lambda: self.browse_xlsx_file(None))
        
        # Bind a keyboard shortcut to the browse button.
        self.master.bind("<Control-e>", lambda event: self.browse_xlsx_file(event))
        self.answer_table_var = BooleanVar()
        self.answer_table_checkbutton = Checkbutton(
            self.input_frame, text="Answer table format", variable=self.answer_table_var)

        self.docx_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.docx_entry.grid(row=1, column=0, columnspan=2,
                             sticky="ew", padx=5, pady=5)
        self.docx_button.grid(row=1, column=2, padx=5, pady=5)

        self.xlsx_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.xlsx_entry.grid(row=3, column=0, columnspan=2,
                             sticky="ew", padx=5, pady=5)
        self.xlsx_button.grid(row=3, column=2, padx=5, pady=5)

        self.answer_table_checkbutton.grid(
            row=4, column=0, columnspan=3, sticky="w", padx=5, pady=5)

        # Create button widgets
        self.convert_button = Button(
            self.button_frame, text="Convert", command=lambda: self.convert(None))
        
        # Bind a keyboard shortcut to the browse button.
        self.master.bind("<Return>", lambda event: self.convert(event))
        self.reset_button = Button(
            self.button_frame, text="Reset", command=lambda: self.reset(None))
        
        # Bind a keyboard shortcut to the browse button.
        self.master.bind(
            "<Control-r>", lambda event: self.reset(event))
        self.result_label = ScrolledText(
            self.button_frame, width=60, height=10, wrap=WORD)
        self.convert_button.grid(row=0, column=0, padx=5, pady=5)
        self.reset_button.grid(row=0, column=1, padx=5, pady=5)

        # Configure result label widget
        self.result_label.grid(row=1, column=0, columnspan=2,
                               padx=5, pady=5, sticky="nsew")
        self.result_label.config(state=DISABLED)

        # Configure grid layout
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(0, weight=1)
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_columnconfigure(1, weight=1)
        self.button_frame.grid_rowconfigure(1, weight=1)

    def browse_docx_file(self, event):
        """
        Opens a file dialog to allow the user to select a docx file.

        Args:
            event: The event that triggered the function.

        Returns:
            The path to the selected file.
        """

        filename = filedialog.askopenfilename(
            filetypes=[("Document files", "*.docx *.doc")], initialdir=os.getcwd())
        self.docx_entry.delete(0, END)
        self.docx_entry.insert(0, filename)

    def browse_xlsx_file(self, event):
        """
        Opens a file dialog to allow the user to select an xlsx file.

        Args:
            event: The event that triggered the function.

        Returns:
            The path to the selected file.
        """
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")], initialdir=os.getcwd())
        self.xlsx_entry.delete(0, END)
        self.xlsx_entry.insert(0, filename)

    def convert(self, event):
        """
        Convert a docx file to an xlsx file.

        Args:
            event: The event that triggered the function.
        """
        docx_filename = self.docx_entry.get()
        xlsx_filename = self.xlsx_entry.get()
        answer_table = self.answer_table_var.get()
        if not docx_filename or not xlsx_filename:
            self.result_label.config(state=NORMAL)
            self.result_label.delete('1.0', END)
            self.result_label.insert(
                END, "Please select both a docx file and an xlsx file.")
            self.result_label.config(state=DISABLED)
            return
        if not docx_filename.endswith(".docx") or not xlsx_filename.endswith(".xlsx"):
            self.result_label.config(state=NORMAL)
            self.result_label.delete('1.0', END)
            self.result_label.insert(END, "Wrong file format.")
            self.result_label.config(state=DISABLED)
            return
        converter_obj = converter.QuizConverter(
            docx_filename=docx_filename, excel_filename=xlsx_filename, answer_table=answer_table)
        converter_obj.convert()
        self.result_label.config(state=NORMAL)
        self.result_label.delete('1.0', END)
        self.result_label.insert(END, "Saved completed successfully!\nLogs:\n")
        logs = Logs().read()
        if logs:
            self.result_label.insert(END, logs)
        else:
            self.result_label.insert(END, "No logs found.")
        self.result_label.config(state=DISABLED)

    def reset(self, event):
        """
        reset the file paths to empty strings.

        args:
            event: the event that triggered the function.
        """
        self.docx_entry.delete(0,END)
        self.xlsx_entry.delete(0, END)
        self.result_label.config(state=NORMAL)
        self.result_label.delete('1.0',END)
        self.result_label.insert(END, "no logs found.")
        self.result_label.config(state=DISABLED)

    def open_browser(sefl, url):
        """
        Reset the file paths to empty strings.

        Args:
            event: The event that triggered the function.
        """
        open_new_tab(url=url)
