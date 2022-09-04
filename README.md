# CBSE Result TXT to Excel Converter and Analyzer
A GUI and CLI tool to convert the Class 10th and 12th Result Text File to Excel with proper formatting and analyze the result and display some statistics.

## Features
* Simple and easy to use. Single Click
* Converts the txt file into properly formatted excel file.
* Different Spreadsheet page for each Subject
* Displays the statistics such as 
  * Top 5 Male and Female Students
  * Number of Distinctions
  * Number of Distinctions in all 5 subjects
  * Children with full marks in individual subjects

## How to Run

### There are two methods to run
1. Google Colab (Easiest)
2. Locally using python

### Google Colab
1. Go to the link (https://colab.research.google.com/drive/1ardBfRG_S40qejG5VnCVWIJYEGpSHluI?usp=sharing)
2. Run each cell one by one
3. Rest of the instructions are given in the colab file

### Locally using Python
1. Install [Python](https://www.python.org/downloads/) (if not already)
    * While Installing make sure to check the add to system PATH option
2. Download the code by clicking on the Green Code button then Download as ZIP
    * ![image](https://user-images.githubusercontent.com/87975651/188325450-7c2e950a-cd7a-4d07-b9c2-5f73a4e177a4.png)
4. Extract the ZIP file and paste your result text file in the extracted folder
5. Open the terminal in the extracted folder and run `pip install -r requirements.txt`
6. Run `main.py`
    * Enter the file name
    * Enter the output file name.. make sure to enter .xlsx in after the file name
    * Enter the mode i.e the class. It can accept only two values i.e. `12th` and `10th`
    * Exported File will automatically Launch.
  
## You can run the program again for another class
