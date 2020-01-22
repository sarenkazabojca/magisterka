import xlwt

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Sheet_1')
workbook.save('summary.xls')
