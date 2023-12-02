# This script runs the quiz converter program
import argparse
from module.write_new_file import WRITE_NEW_EXCEL
from utils import converter
import os


if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog="main.py")
    parser.add_argument("-d", "--docx", help="filename.docx path(s)")
    # parser.add_argument("-x", "--xlsx", default="./template.xlsx", help="filename.xlsx path(s) [default: template.xlsx]")
    parser.add_argument("--answer-table", default=False, action=argparse.BooleanOptionalAction, help="--annswer-table if your file has answer table format")
    parser.add_argument("--is-underline", default=False, action=argparse.BooleanOptionalAction, help="--is-underline if your file answer has answer underline format")
    parser.add_argument("--is-italic", default=False, action=argparse.BooleanOptionalAction, help="--is-italic if your file answer has answer italic format")
    args = parser.parse_args()
    # Check input user
    if not args.docx:
        print("Not found file docx. Pls check your input!!!")
        exit(1)
    if not args.docx.endswith(".docx"):
        print("wrong format file")
        exit(1)
    path_xlsx = WRITE_NEW_EXCEL(args.docx)
    # Create a QuizConverter object with the input and output filenames
    converter = converter.QuizConverter(
        docx_filename=args.docx, 
        excel_filename=path_xlsx, 
        answer_table=bool(args.answer_table), 
        is_underline=bool(args.is_underline),
        is_italic=bool(args.is_italic))
    # Call the convert method to extract data from the docx file and write it to the excel file
    converter.convert()
