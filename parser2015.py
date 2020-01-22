import xlrd
import xlwt
import re
import os

summaryWorkbook = xlwt.Workbook()
summarySheet = summaryWorkbook.add_sheet('dane')
currentRow = 0

dirname = os.getcwd() + "\\data\\2015"

for filename in os.listdir(dirname):
    path = os.path.join(dirname, filename)
    workbook = xlrd.open_workbook(path)
    print (">> Processing workbook: " + filename)

    for i in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(0)
        river = sheet.cell(0, 1).value
        print (">> Processing sheet: \"" + sheet.name + " : " + river + "\" [" + str(sheet.nrows) + "]")

        if currentRow == 0:
            summarySheet.write(0, 0, "Nazwa ppk")
            summarySheet.write(0, 1, "Data")
            summarySheet.write(0, 2, sheet.cell(15, 8).value)
            summarySheet.write(0, 3, sheet.cell(15, 13).value)
            summarySheet.write(0, 4, sheet.cell(15, 14).value)
            summarySheet.write(0, 5, sheet.cell(15, 16).value)
            summarySheet.write(0, 6, sheet.cell(15, 20).value)
            summarySheet.write(0, 7, sheet.cell(15, 26).value)
            summarySheet.write(0, 8, sheet.cell(15, 27).value)
            summarySheet.write(0, 9, sheet.cell(15, 29).value)
            summarySheet.write(0, 10, sheet.cell(15, 30).value)
            summarySheet.write(0, 11, sheet.cell(15, 31).value)
            summarySheet.write(0, 12, sheet.cell(15, 33).value)
            summarySheet.write(0, 13, sheet.cell(15, 34).value)
            summarySheet.write(0, 14, sheet.cell(15, 35).value)
            summarySheet.write(0, 15, sheet.cell(15, 90).value)
            summarySheet.write(0, 16, sheet.cell(15, 91).value)
            currentRow = 1

        for i in range(sheet.nrows):
            if re.search("[\\n\\s]*\d{2}[\.-]\d{2}[\.-]\d{4}[\\n\\s]*", str(sheet.cell(i, 0).value)):
                print ("Writing: " + str(sheet.cell(i, 0).value))
                summarySheet.write(currentRow, 0, river)
                summarySheet.write(currentRow, 1, sheet.cell(i,0).value)
                summarySheet.write(currentRow, 2, sheet.cell(i, 8).value)
                summarySheet.write(currentRow, 3, sheet.cell(i, 13).value)
                summarySheet.write(currentRow, 4, sheet.cell(i, 14).value)
                summarySheet.write(currentRow, 5, sheet.cell(i, 16).value)
                summarySheet.write(currentRow, 6, sheet.cell(i, 20).value)
                summarySheet.write(currentRow, 7, sheet.cell(i, 26).value)
                summarySheet.write(currentRow, 8, sheet.cell(i, 27).value)
                summarySheet.write(currentRow, 9, sheet.cell(i, 29).value)
                summarySheet.write(currentRow, 10, sheet.cell(i, 30).value)
                summarySheet.write(currentRow, 11, sheet.cell(i, 31).value)
                summarySheet.write(currentRow, 12, sheet.cell(i, 33).value)
                summarySheet.write(currentRow, 13, sheet.cell(i, 34).value)
                summarySheet.write(currentRow, 14, sheet.cell(i, 35).value)
                summarySheet.write(currentRow, 15, sheet.cell(i, 90).value)
                summarySheet.write(currentRow, 16, sheet.cell(i, 91).value)
                currentRow = currentRow + 1

summaryWorkbook.save(os.getcwd() + "\\summary2015.xls")
print ("Finished")