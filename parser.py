import xlrd
import xlwt
import re
import os

cWorkbook = xlwt.Workbook()
cSheet = cWorkbook.add_sheet('Sheet_1')
cWorkbook.save('summary.xls')

dirname = os.getcwd() + "\\..\\data"

for filename in os.listdir(dirname):
    path = os.path.join(dirname, filename)
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    print (sheet.name)
    print (sheet.cell(0, 1).value)

    for i in range(sheet.nrows):
        if re.search("\\d{2}-\d{2}-\d{4}", str(sheet.cell(i,0).value)):
            print (sheet.cell(i,0).value, sheet.cell(i,8).value)

    #if sheet.cell(8, 18).value == xlrd.empty_cell.value:
    #   print ('empty')