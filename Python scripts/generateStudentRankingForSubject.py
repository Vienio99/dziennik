import os, openpyxl

os.system('cls')

def generateStudentRankingForSubject(subject):

  wb = openpyxl.load_workbook('oceny-grupa1.xlsx')
  sheet = wb[subject]

  ranking = {}

  for i in range(1, sheet.max_row + 1):
    if sheet.cell(row=i, column=2).value in ranking:
      ranking[sheet.cell(row=i, column=2).value].append(sheet.cell(row=i, column=1).value)
    else:
      ranking[sheet.cell(row=i, column=2).value] = [sheet.cell(row=i, column=1).value]

  ranking = dict(sorted(ranking.items(), reverse=True))

  counter = 1
  for students in ranking.values():
    position = ''
    for student in students:
      position += student + ', '
    print(str(counter) + '. ' + position[:-2])
    counter += 1
    
subject = input('Enter a name of the subject: ')
generateStudentRankingForSubject(subject)