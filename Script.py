import sys
import pandas as pd
import openpyxl
import csv
import datetime
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment


#python.exe Script.py export export(1)

#open the two files or one
first_report = 'C:\\Users\\f5mgs\\Downloads\\' + sys.argv[1] + '.csv'
second_report = 'C:\\Users\\f5mgs\\Downloads\\' + sys.argv[2] + '.csv'
filenames = []
if first_report:
    filenames.append(first_report)
if second_report:
    filenames.append(second_report)
outfile = 'C:\\Users\\f5mgs\\Desktop\\Co-op\\Change Meetings\\combined.csv'

#combine the files
a = pd.read_csv(first_report)
b = pd.read_csv(second_report)
frames = [a,b]
result = pd.concat(frames)
result.to_csv(outfile, index=False)


wb = openpyxl.Workbook()
ws = wb.active

with open('C:\\Users\\f5mgs\\Desktop\\Co-op\\Change Meetings\\combined.csv') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        ws.append(row)

new_outfile = 'C:\\Users\\f5mgs\\Desktop\\Co-op\\Change Meetings\\combined.xlsx'
wb.save(new_outfile)

#then open the excel
wb = openpyxl.load_workbook(new_outfile)
sheet = wb.active

#1. delete the rows with change managment date of not today's date

sheet_size = sheet.max_row
today = datetime.datetime.today().date()
shifted = True


for loop in range(20):
    for ro in range(sheet_size + 1):
        if ro == 0 or ro == 1:
            continue
        cm_date = sheet.cell(row=ro, column=9)
        if cm_date.value == None:
            sheet.delete_rows(ro,1)
            continue
        elif len(cm_date.value.split(' ')) != 3:
            sheet.delete_rows(ro,1)
            continue
        cm_date = cm_date.value.split(' ')[2]
        current = datetime.datetime.strptime(cm_date, '%m/%d/%Y').date()
        if current != today:
            sheet.delete_rows(ro,1)

#2 . highlight the rows with change scheuled downtime start or  ad hoc director to true

yelFill = openpyxl.styles.PatternFill(start_color='00FFFF00',
                   end_color='00FFFF00',
                   fill_type='solid')

sheet_size = sheet.max_row
for ro in range(sheet_size + 1):
    if ro == 0 or ro == 1:
        continue
    print sheet.cell(row=ro, column=8).value
    if (sheet.cell(row=ro, column=8).value == 'True') or (sheet.cell(row=ro, column=5).value != None):
        for r in range(sheet.max_column + 1):
            if r == 0:
                continue
            sheet.cell(row=ro, column=r).fill = yelFill



#3. Do borders on every cell and align left
side = Side(border_style='thin', color="FF000000")
column_size = sheet.max_column
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

for i in range(sheet_size + 1):
    if i == 0:
        continue
    for j in range(column_size + 1):
        if j == 0:
            continue
        cell = sheet.cell(row=i, column=j)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='left')

#6. Normalize column widths
sheet.column_dimensions['A'].width = 12
sheet.column_dimensions['B'].width = 50
for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I']:
    sheet.column_dimensions[col].width = 30

#7. Change wording on column header
sheet.cell(row=1, column=8).value = 'Director Approval'

#4. hide the columns that are not needed.
for col in ['J', 'K', 'L']:
    sheet.column_dimensions[col].hidden = True

#save and close the file.
wb.save(new_outfile)