from openpyxl import load_workbook as lw
import unidecode
import re

#Reading the Workbook
wb = lw('US-News Ranking.xlsx', data_only= True)

#Getting sheet names
sheets = wb.sheetnames


#Finds which column contains tuition fee
def tuitionColumn(sheetname, wb):
    w_sh = wb[sheetname]
    letters = ["A", "B", "C", "D","E","F","G", "H","I","J"]
    results = []
    for j in range(1, len(letters)):
        cell = w_sh.cell(row=1,column=j).value
        if cell == None:
            continue
        matched = re.search(r"(^tuition)", cell, re.IGNORECASE)
        if matched != None:
            col = j
            break
        else:
            col = None
            continue
    return col

#Read through University names, Department and Tuition fee
def readSheet(sheetname, wb,tC): #tC = tuition column
    if tC == None:
        return
    i = 2
    w_sh = wb[sheetname]
    lst = []
    while w_sh.cell(row=i, column=tC).value != None:
        university = w_sh.cell(row=i, column=2).value
        university = unidecode.unidecode(university)
        tuition_fee = w_sh.cell(row=i, column=tC).value
        try:
            tuition_fee = float(re.sub(r'[\D]', "", tuition_fee))
        except:
            tuition_fee = ''
        lst.append((university,sheetname,tuition_fee))
        i = i + 1
    return lst

#Main
if __name__ == '__main__':
    for sheet in sheets:
        tc = tuitionColumn(sheet,wb) 
        temp = readSheet(sheet,wb,tc)
        if temp == None:
            continue
        else:
            for i in temp:
                print(i)
        temp.clear()






