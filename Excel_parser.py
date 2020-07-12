import xlrd
import xlsxwriter
import os
import csv

output_filename="output.xlsx"

# Read files
data = {}
for filename in os.listdir("."):
    if filename == output_filename: #ignore output file
        continue
    if filename.endswith(".xls") or filename.endswith(".xlsx"): # Iterate through excel sheets
        print(">>>> Processing xls/xlsx file :{filename}".format(filename=filename))
        data[filename] = {}
        wb = xlrd.open_workbook(filename)
        for sheet in wb.sheets():
            data[filename][sheet.name] = []
            for col_idx in range(0, sheet.ncols):  # Iterate through columns
                cell_value = sheet.cell_value(0, col_idx)  # Read header column value
                data[filename][sheet.name].append(cell_value)
    if filename.endswith(".csv"):
        print(">>>> Processing csv file :{filename}".format(filename=filename))
        data[filename]={}
        with open(filename) as csvfile:
            csv_reader = csv.reader(csvfile, delimiter=',')
            for row in csv_reader:
                data[filename][''] = row
                break #only read first row

#Write output file
workbook = xlsxwriter.Workbook(output_filename)
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
worksheet.write_row(0,0,['File Name','Sheet Name', 'Header Column'], bold)

row=1
for filename in data:
    for sheetname in data[filename]:
        for header_cell in data[filename][sheetname]:
            worksheet.write_row(row, 0, [filename, sheetname, header_cell])
            row += 1

workbook.close()
