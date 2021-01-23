import os, openpyxl

os.system('cls')


def generateStudentRankingForAllSubjects():

  wb = openpyxl.load_workbook('oceny-grupa1.xlsx')

  grades = {}
  for sheet in wb.worksheets:
    sheet = wb[sheet.title]
    for i in range(1, sheet.max_row + 1):
      if sheet.cell(row=i, column=1).value in grades:
        grades[sheet.cell(row=i, column=1).value].append(sheet.cell(row=i, column=2).value)
      else:
        grades[sheet.cell(row=i, column=1).value] = [sheet.cell(row=i, column=2).value]

  for student in grades.keys():
    grades[student] = round(sum(grades[student])/len(wb.worksheets), 2)

  ranking = {}

  for student, average in grades.items():
    if average in ranking:
      ranking[average].append(student)
    else:
      ranking[average] = [student]

  ranking = dict(sorted(ranking.items(), reverse=True))

  counter = 1

  for students in ranking.values():
    position = ''
    for student in students:
      position += student + ', '
    
    print(str(counter) + '. ' + position[:-2])
    counter += 1


generateStudentRankingForAllSubjects()