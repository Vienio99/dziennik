import os, openpyxl
import pandas as pd
import matplotlib

os.system('cls')

def calculateAverageForAllSubjects():

  wb = openpyxl.load_workbook('oceny-grupa1.xlsx')

  averages = {}

  for sheet in wb.worksheets:
    sheet = wb[sheet.title]
    average = []

    for i in range(1, sheet.max_row + 1):
        average.append(sheet.cell(row=i, column=2).value)

    averages[sheet.title] = round(sum(average) / sheet.max_row, 2)

  averages = sorted([(value,key) for (key,value) in averages.items()], reverse=True)

  subjectsSeparate = []
  averagesSeparate = []
  for x in range(len(averages)):
    averagesSeparate.append(averages[x][0])
    subjectsSeparate.append(averages[x][1])


  df = pd.DataFrame({'Subject':subjectsSeparate, 'Average':averagesSeparate})
  ax = df.plot.bar(x='Subject', y='Average', rot=0)

calculateAverageForAllSubjects()