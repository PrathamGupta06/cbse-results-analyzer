# CBSE Result TXT to Excel Converter and Analyzer
A GUI and CLI tool to convert the Class 10th and 12th Result Text File to Excel with proper formatting and analyze the result and display some statistics.
Currently in Development. Clone the [Dev](https://github.com/PrathamGupta06/cbse-results-analyzer/tree/dev) branch.

## How to Run

Install [Python 3](https://www.python.org/downloads/)

### Modules Required
```
pip install pandas
pip install openpyxl
pip install XlsxWriter
```
In the main.py file
Change the below variables
```py
input_file = r'input/result_10th.txt' #location of the input file
output_path_excel = r'output/result10th.xlsx' #location of the excel file to be saved to. The file will be overwritten if already exists.
mode = '10th'
# mode = '12th' for 12th class
```
And run the program. It will automatically open the Excel File.
