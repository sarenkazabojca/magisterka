import xlrd
import xlwt
import re
import os

#tworzymy nowy plik z arkuszem o nazwie "dane"
summaryWorkbook = xlwt.Workbook()
summarySheet = summaryWorkbook.add_sheet('dane')
currentRow = 0

#lokalizacja danych
dirname = os.getcwd() + "\\data\\2014"

#pętla przechodząca kolejno po plikach z danymi
for filename in os.listdir(dirname):
    path = os.path.join(dirname, filename)
    workbook = xlrd.open_workbook(path)
    print (">> Processing workbook: " + filename)

#obsługa arkuszy w plikach
    for i in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(i)

        firstcell = sheet.cell(0, 0).value
        start = 0
        if firstcell is None or firstcell == "":
            start = 1
#nazwa z komórki (start,1), a jeśli ta jest pusta, z komórki (start, 2)
        rivername = sheet.cell(start, 1).value or sheet.cell(start, 2).value
        print (">> Processing sheet: \"" + sheet.name + " : " + rivername + "\" [" + str(sheet.nrows) + "]")

# nadajemy nazwy kolumnom w nowym pliku
        if currentRow == 0:
            summarySheet.write(start, 0, "Nazwa ppk")
            summarySheet.write(start, 1, "Stat")
            summarySheet.write(start, 2, "Temperatura (oC)")
            summarySheet.write(start, 3, "Barwa (mg/l Pt)")
            summarySheet.write(start, 4, "Zawiesina ogólna(mg / l)")
            summarySheet.write(start, 5, "Tlen rozpuszczony (mg O2/l)")
            summarySheet.write(start, 6, "BZT5 (mg O2/l)")
            summarySheet.write(start, 7, "OWO (mg C/l)")
            summarySheet.write(start, 8, "Przewodność w 20oC (uS/cm)")
            summarySheet.write(start, 9, "Substancje rozpuszczone (mg/l)")
            summarySheet.write(start, 10, "Siarczany (mg SO4/l)")
            summarySheet.write(start, 11, "Chlorki (mg Cl/l)")
            summarySheet.write(start, 12, "Wapń (mg Ca/l)")
            summarySheet.write(start, 13, "Magnez (mg Mg/l)")
            summarySheet.write(start, 14, "Twardość ogólna (mg CaCO3/l)")
            summarySheet.write(start, 15, "Odczyn pH")
            summarySheet.write(start, 16, "Zasadowość ogółna (mg CaCO3/l)")
            summarySheet.write(start, 17, "Azot amonowy (mg N-NH4/l)")
            summarySheet.write(start, 18, "Azot Kjeldahla (mg N/l)")
            summarySheet.write(start, 19, "Azot azotanowy (mg N-NO3/l)")
            summarySheet.write(start, 20, "Azot ogólny (mg N/l)")
            summarySheet.write(start, 21, "Fosforany  (mg PO4/l)")
            summarySheet.write(start, 22, "Fosfor ogólny (mg P/l)")
            currentRow = 1

# w danych wyznaczamy komórki do spisania i przepisujemy je do pliku wyjściowego
# dane spisujemy tylko jeśli odpowiednio dla danego wiersza w pierszej kolumnie znajduje się napis Średnia, Max lub Min
        for i in range(sheet.nrows):
            if re.search("(Średnia|Max|Min)", str(sheet.cell(i, 0).value)):
                print ("Writing: " + str(sheet.cell(i, 0).value))
                summarySheet.write(currentRow, 0, rivername)
                summarySheet.write(currentRow, 1, sheet.cell(i, 0).value)
                summarySheet.write(currentRow, 2, sheet.cell(i, 7).value)
                summarySheet.write(currentRow, 3, sheet.cell(i, 9).value)
                summarySheet.write(currentRow, 4, sheet.cell(i, 11).value)
                summarySheet.write(currentRow, 5, sheet.cell(i, 12).value)
                summarySheet.write(currentRow, 6, sheet.cell(i, 13).value)
                summarySheet.write(currentRow, 7, sheet.cell(i, 15).value)
                summarySheet.write(currentRow, 8, sheet.cell(i, 19).value)
                summarySheet.write(currentRow, 9, sheet.cell(i, 20).value)
                summarySheet.write(currentRow, 10, sheet.cell(i, 21).value)
                summarySheet.write(currentRow, 11, sheet.cell(i, 22).value)
                summarySheet.write(currentRow, 12, sheet.cell(i, 23).value)
                summarySheet.write(currentRow, 13, sheet.cell(i, 24).value)
                summarySheet.write(currentRow, 14, sheet.cell(i, 25).value)
                summarySheet.write(currentRow, 15, sheet.cell(i, 26).value)
                summarySheet.write(currentRow, 16, sheet.cell(i, 27).value)
                summarySheet.write(currentRow, 17, sheet.cell(i, 28).value)
                summarySheet.write(currentRow, 18, sheet.cell(i, 29).value)
                summarySheet.write(currentRow, 19, sheet.cell(i, 30).value)
                summarySheet.write(currentRow, 20, sheet.cell(i, 32).value)
                summarySheet.write(currentRow, 21, sheet.cell(i, 33).value)
                summarySheet.write(currentRow, 22, sheet.cell(i, 34).value)
                currentRow = currentRow + 1
            else:
                print ("Skipping: " + str(sheet.cell(i, 0).value) + " (" + str(sheet.cell(i, 0).ctype) + ")")

#zapisujemy zmiany do pliku wyjściowego
summaryWorkbook.save(os.getcwd() + "\\summary2011-14.xls")
print ("Finished")