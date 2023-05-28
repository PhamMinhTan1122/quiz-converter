# **Quiz Converter**

## **Description**

Quiz Converter is a Python program that converts quiz questions from docx format to excel format. It can be used by teachers or students who want to create or take quizzes online using platforms like Quizizz.

## **🎇Installation**

You also need to clone this repository or download the zip file:

```bash 
git clone https://github.com/PhamMinhTan1122/quizizz-convert.git
```


To install and run Quiz Converter, you need to have Python 3 and the following libraries installed on your system:

- openpyxl
- python-docx

You can install these libraries using pip:

```bash
pip install openpyxl
pip install python-docx
```
or

```bash
pip install -r requirements.txt
```

## **🕹Usage**
To use Quiz Converter, you need to have a docx file that contains the quiz questions and their options. The questions should be numbered and the correct option should be underlined. For example:

![alt text](https://raw.githubusercontent.com/PhamMinhTan1122/quizizz-convert/main/public/imgs/data-raw.png "Data-raw")

Then, you need to run the main.py script from the scripts directory and provide the input and output filenames as arguments. For example:

```bash
python main.py -d <your file docx (Word)> -x <your file xlsx (File SpreadSheets of Quizziz)>
```

or

```bash
python main.py -docx <your file docx (Word)> -xlsx <your file xlsx (File SpreadSheets of Quizziz)>
```

### 🚨Noticed!

You can download SpreadSheets file on Quizizz:

![alt text](https://raw.githubusercontent.com/PhamMinhTan1122/quizizz-convert/main/public/imgs/download-spreadsheets.png "SpreadSheets file")

or

You can own make file in excel or Google Sheets (if you use Google Sheets, you can download .xlsx format File -> Download -> Microsoft Excel (.xlsx)), For example:

![alt text](https://raw.githubusercontent.com/PhamMinhTan1122/quizizz-convert/main/public/imgs/excel_before.png "Excel before")

This will create an excel file that contains the quiz questions and their options in separate columns. The correct option will also be marked in column G. For example:

![alt text](https://raw.githubusercontent.com/PhamMinhTan1122/quizizz-convert/main/public/imgs/excel_after.png "Excel after")
After run and sucessful, You can then upload this excel file to Quizizz or any other online quiz platform that supports this format.

## TODO
- [X] ~~Get answer format underline~~
- [X] ~~Get answer format table~~

## License
Quiz Converter is licensed under the MIT License. See [LICENSE.txt](https://raw.githubusercontent.com/PhamMinhTan1122/quizizz-convert/main/docs/LICENSE.txt) for more details.

## Contributing
Quiz Converter is an open source project and welcomes contributions from anyone. If you want to contribute to this project, please follow these steps:

- Fork this repository and clone it to your local machine.
- Create a new branch for your feature or bug fix.
- Make your changes and commit them with descriptive messages.
- Push your branch to your forked repository and create a pull request.
- Wait for your pull request to be reviewed and merged.
- Please follow the PEP 8 style guide for Python code and write clear and concise comments.

## Contact
If you have any questions or feedback about Quiz Converter, please feel free to contact me at:

- Email: masterminhtan@gmail.com
- Twitter: @masterminhtan