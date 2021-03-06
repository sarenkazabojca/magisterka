import xlrd
import xlwt
import re
import os

summaryWorkbook = xlwt.Workbook()
summarySheet = summaryWorkbook.add_sheet('dane')
currentRow = 0

dirname = os.getcwd() + "\\data\\2017"

for filename in os.listdir(dirname):
    path = os.path.join(dirname, filename)
    workbook = xlrd.open_workbook(path)
    print ("Processing workbook: " + filename)

    for i in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(0)
        river = sheet.cell(0, 1).value
        print ("Processing sheet: \"" + sheet.name + " : " + river + "\"")

        if currentRow == 0:
            summarySheet.write(0, 0, "Nazwa ppk")
            summarySheet.write(0, 1, "Data")
            summarySheet.write(0, 2, "Temperatura (oC)")
            summarySheet.write(0, 3, "Tlen rozpuszczony (mg O2/l)")
            summarySheet.write(0, 4, "BZT5 (mg O2/l)")
            summarySheet.write(0, 5, "OWO (mg C/l)")
            summarySheet.write(0, 6, "Przewodność w 20oC (uS/cm)")
            summarySheet.write(0, 7, "Twardość ogólna (mg CaCO3/l)")
            summarySheet.write(0, 8, "Odczyn pH")
            summarySheet.write(0, 9, "Azot amonowy (mg N-NH4/l)")
            summarySheet.write(0, 10, "Azot Kjeldahla (mg N/l)")
            summarySheet.write(0, 11, "Azot azotanowy (mg N-NO3/l)")
            summarySheet.write(0, 12, "Azot ogólny (mg N/l)")
            summarySheet.write(0, 13, "Fosforany  (mg PO4/l)")
            summarySheet.write(0, 14, "Fosfor ogólny (mg P/l)")
            summarySheet.write(0, 15, "Benzo(k)fluoranten (µg/l)")
            summarySheet.write(0, 16, "Benzo(g,h,i)perylen (µg/l)")
            currentRow = 1

#TODO
        for i in range(sheet.nrows):
            if re.search("\\d{2}-\d{2}-\d{4}", str(sheet.cell(i,0).value)):
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

summaryWorkbook.save(os.getcwd() + "\\summary2017.xls")
print ("Finished")