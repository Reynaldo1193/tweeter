# Function to count frequency of each element  
import collections 
import xlsxwriter
import xlrd
import sys
import re

def myFunc(e):
  return e['total']

def freq(str):
    str2=[]
    temasF = []    
    
    for i in range(0, len(str)):              
        str[i] = str[i].replace(",","")
        str[i] = str[i].replace(" ","")
        if str[i] not in str2: 
              
            str2.append(str[i])
            print(str[i])
            
    for i in range(0, len(str2)): 

        temasF.append({"palabra": str2[i], "total": str.count(str2[i])})

    temasF.sort(key=myFunc, reverse=True)
    print(len(temasF))
    """ workbook = xlsxwriter.Workbook("TopDeTemas.xlsx")
    worksheet = workbook.add_worksheet()     """

    for i in range(0, len(temasF)):
        #worksheet.write_string(i, 0, temasF[i]["palabra"])
        #worksheet.write_number(i, 1, temasF[i]["total"])        
        
        print(i)
        print(temasF[i])

    #workbook.close()
    #print("Excel file ready") 
                
    
def freqreg(str):
    tot = 0
    pos = 0
    
    for i in range(0, len(str)):              

        if re.search("gobernadores", str[i].lower()):
            print(str[i])
            print(pos)
            print("====================================")
            tot = tot +1

        pos = pos +1

    print(tot)



temas = []
index = 0

workbook = xlrd.open_workbook("análisisTwitter.xlsx")
sheet = workbook.sheet_by_index(0)

for i in range(sheet.nrows):
    temas.append(sheet.cell_value(i, 8))

freq(temas)