# This script runs the quiz converter program
import argparse
from utils import converter

if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog="main.py")
    parser.add_argument("-d", "--docx", help="filename.docx path(s)")
    parser.add_argument("-x", "--xlsx", default="./QuizizzSampleSpreadsheetUpdated.xlsx", help="filename.xlsx path(s)")
    parser.add_argument("--answer-table", default=False, action=argparse.BooleanOptionalAction, help="--annswer-table if your file has answer table format")
    args = parser.parse_args()
    # Check input user
    if not args.docx or not args.xlsx:
        print("Not found file docx or xlsx. Pls check your input!!!")
        exit(1)
    if not args.docx.endswith(".docx") or not args.xlsx.endswith(".xlsx"):
        print("wrong format file")
        exit(1)
    # Create a QuizConverter object with the input and output filenames
    converter = converter.QuizConverter(docx_filename=args.docx, excel_filename=args.xlsx, answer_table=bool(args.answer_table))
    # Call the convert method to extract data from the docx file and write it to the excel file
    converter.convert()
  
