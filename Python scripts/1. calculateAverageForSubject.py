import os, openpyxl

os.system('cls')



def calculateAverageForSubject(subject):
    wb = openpyxl.load_workbook('oceny-grupa1.xlsx')
    sheet = wb[subject]

    gradesList = []
    for i in range(1, sheet.max_row + 1):
        gradesList.append(sheet.cell(row=i, column=2).value)
    
    average = sum(gradesList) / len(gradesList)
    print(round(average, 2))

subject = input('Enter a name of the subject: ')
calculateAverageForSubject(subject)




