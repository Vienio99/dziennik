import os, openpyxl

os.system('cls')

def calculateAverageForStudent(student):
    wb = openpyxl.load_workbook('oceny-grupa1.xlsx')
    gradesList = []
    for sheet in wb.worksheets:
        sheet = wb[sheet.title]
        for i in range(1, sheet.max_row + 1):
            if sheet.cell(row=i, column=1).value == student:
                gradesList.append(sheet.cell(row=i, column=2).value)
    average = sum(gradesList) / len(gradesList)
    print(average)

student = input('Enter a name of the student: ')
calculateAverageForStudent(student)