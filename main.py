"""
MADE BY PRATHAM GUPTA (https://github.com/PrathamGupta06/cbse-results-analyzer)
"""

import os
import sys
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from _functions import *

# TODO:
#     Fix for case where student is absent

"""Change the Below Variables"""

# Values for testing
# input_file = r'input/12th.txt'
# output_path_excel = r'output/12th.xlsx'
# mode = '12th'

print("Welcome to CBSE Results Analyzer.\nThis Program is made by Pratham Gupta ("
      "https://github.com/PrathamGupta06/cbse-results-analyzer)\n")

if len(sys.argv) > 1:
    # Command Line Argument Mode
    input_file = sys.argv[1]
    output_path_excel = sys.argv[2]
    mode = sys.argv[3]
else:
    # Taking input from user
    input_file = input("Enter Input Result File Path (e.g 12th.txt): ")
    output_path_excel = input("Enter Output Path (e.g 12th.xlsx): ")
    mode = input("Enter Class (10th or 12th): ")

# Create Workbook
wb = Workbook()
save_wb(wb, output_path_excel)
print("Blank Workbook Created")
wb.close()

column_widths = {
    "Roll No": 20,
    "Gender": 8,
    "Name": 20,
    "Subject": 15,
    "Grade": 6,
    "GR1": 5,
    "GR2": 5,
    "GR3": 5,
    "Result": 8,
    "Best 5 Percentage": 10,
    "Compartment Subject": 15
}

# Get the data from the file
headers, data, AllSubjectNames = get_data(input_file, mode)
main_df = pd.DataFrame(data, columns=headers)

# Convert Pandas dataframe to Excel
wb = load_workbook(output_path_excel)
main_sheet = wb["Sheet"]
main_sheet.title = "Main Result Sheet"
write_data(main_sheet, main_df)
adjust_column_widths(main_sheet, column_widths)

# Freezing the First Row and First Column of Main Sheet
main_sheet.freeze_panes = "D2"

# Creating Individual Subject Sheets
for SubjectName in AllSubjectNames:
    index_no = main_df.columns.get_loc(SubjectName)
    subject_df = main_df.iloc[:, [0, 1, 2, index_no, index_no + 1]]  # filter the columns to the subject name and marks
    subject_df = subject_df[subject_df[SubjectName] >= 0]  # filter the rows which do not have subject marks
    subject_sheet = wb.create_sheet(SubjectName)
    write_data(subject_sheet, subject_df)
    adjust_column_widths(subject_sheet, column_widths)

# Analyzing the Results

# Top 5 Analyze
# Splitting the data into Male and Female
male = [student for student in data if student[1] == "M"]
female = [student for student in data if student[1] == "F"]

# Sorting the Male and Female list by Best 5 Percentage
top_n_children = 5
best_5_index = headers.index("Best 5 percentage")
top_male = sorted(male, key=lambda student: student[best_5_index], reverse=True)[:top_n_children]
top_female = sorted(female, key=lambda student: student[best_5_index], reverse=True)[:top_n_children]

# Removing the Unnecessary Columns in the list
top_male = [[student[0], student[1], student[2], student[best_5_index]] for student in top_male]
top_female = [[student[0], student[1], student[2], student[best_5_index]] for student in top_female]

# Analyzation related to marks
children_with_full_marks = [["Roll No", "Gender", "Name", "Subject"]]
total_distinctions = 0
all_5_distinctions = 0  # All 5 Subjects Distinction Analyze
all_5_distinction_students = []  # All 5 Subjects Distinction Student Details

for student_data in data:
    student_distinctions = 0  # Individual Student Distinctions
    for index, value in enumerate(student_data):
        if type(value) is int or type(value) is float:
            if value >= 75:
                student_distinctions += 1
                total_distinctions += 1
            if value == 100:
                children_with_full_marks.append([student_data[0], student_data[1], student_data[2], headers[index]])
    #  For All 5 Distinctions
    if student_distinctions >= 5:
        all_5_distinctions += 1
        all_5_distinction_students.append([student_data[0], student_data[1], student_data[2]])

# Writing the analysis data to worksheet
wb.create_sheet('Analysis', 1)
analyze_ws = wb['Analysis']
last_row = 0

# Children with Full Marks
analyze_ws.append(["Children With Full Marks"])
last_row += 1
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=4)

for i in children_with_full_marks:
    analyze_ws.append(i)
    last_row += 1

# Overall Toppers
analyze_ws.append([])
last_row += 1

analyze_ws.append(["Overall {} Toppers".format(top_n_children)])
last_row += 1
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=4)

# Male Toppers
analyze_ws.append(["Male"])
last_row += 1
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=4)

analyze_ws.append(["Roll No", "Gender", "Name", "Best 5 Percentage"])
last_row += 1

for i in top_male:
    analyze_ws.append(i)
last_row += len(top_male)

# Female Toppers
analyze_ws.append([])
last_row += 1

analyze_ws.append(["Female"])
last_row += 1
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=4)

analyze_ws.append(["Roll No", "Gender", "Name", "Best 5 Percentage"])
last_row += 1
for i in top_female:
    analyze_ws.append(i)
last_row += len(top_male)

# Distinctions in all subjects

analyze_ws.append([])
last_row += 1

analyze_ws.append(["Total Distinctions: {0}".format(total_distinctions)])
last_row += 1
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=3)
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)

# Distinctions in all 5 Subjects

analyze_ws.append(["Total Distinctions in all 5 subjects: {}".format(all_5_distinctions)])
last_row += 1
analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=3)
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)

analyze_ws.append([])
last_row += 1

analyze_ws.append(["Distinctions in all 5 Subjects Students"])
last_row += 1
analyze_ws.cell(row=last_row, column=1).font = Font(bold=True)

analyze_ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=3)
for i in all_5_distinction_students:
    analyze_ws.append(i)
last_row += len(all_5_distinction_students)

# Adjust the widths
analyze_ws.column_dimensions["A"].width = column_widths["Roll No"]
analyze_ws.column_dimensions["B"].width = column_widths["Gender"]
analyze_ws.column_dimensions["C"].width = column_widths["Name"]
analyze_ws.column_dimensions["D"].width = column_widths["Subject"]

# Save Workbook
save_wb(wb, output_path_excel)
wb.close()
print("Workbook Saved at", os.path.abspath(output_path_excel))

# Launch Workbook if mode is not Command Line Argument
if len(sys.argv) == 1:
    os.startfile(os.path.abspath(output_path_excel))
    print("Program Completed Successfully")
