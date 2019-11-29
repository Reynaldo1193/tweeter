import xlsxwriter
import xlrd
import sys



workbook = xlrd.open_workbook("análisisTwiter.xlsx")
sheet = workbook.sheet_by_index(0)

for i in range(sheet.nrows):    
    ids.append(sheet.cell_value(i, 0)) 
    fechas.append(sheet.cell_value(i, 1))
    tweets.append(sheet.cell_value(i, 2))
    contadorFav.append(sheet.cell_value(i, 4))
    contadorRT.append(sheet.cell_value(i, 5))
    indexs.append(index)
    index = index + 1
