import openpyxl
import json
from docx import Document

wb = openpyxl.load_workbook(filename = 
'C:\\Users\\shave\\Documents\\Python soubory\\Programy\\Petra excel to word\\NEW2\\BOM.xlsx')
b = wb.worksheets[0]

BOM = list()
for row in range (221,516):
    print(b[f"A{row}"].value)
    a = dict()
    a["typ"]=b[f"A{row}"].value
    a["oil"]=b[f"B{row}"].value
    a["procent"]=b[f"C{row}"].value
    a["terpen"]=b[f"D{row}"].value
    a["typ_o"]=float(b[f"E{row}"].value)
    a["typ_w"]= format(a["typ_o"]*0.05, "0.3f")
    a["oil_o"]=float(b[f"G{row}"].value)
    a["oil_w"]=format(a["oil_o"]*0.05, "0.3f")
    if b[f"I{row}"].value == "N/A":
        a["terpen_o"]=None
        a["terpen_w"]=None
    else:
        a["terpen_o"]=float(b[f"I{row}"].value)
        a["terpen_w"]=format(a["terpen_o"]*0.05, "0.3f")
    if b[f"K{row}"].value == "N/A" or b[f"K{row}"].value == None:
        a["cbd_crude_o"]=None
        a["cbd_crude_w"]=None
    else:
        a["cbd_crude_o"]=float(b[f"L{row}"].value)
        a["cbd_crude_w"]=format(a["cbd_crude_o"]*0.05, "0.3f")
    if b[f"N{row}"].value == None:
        a["cbg_o"]=None
        a["cbg_w"]=None
    else:
        a["cbg_o"]=float(b[f"O{row}"].value)
        a["cbg_w"]=format(a["cbg_o"]*0.05, "0.3f")
    if b[f"Q{row}"].value == None:
        a["mct_o"]=None
        a["mct_w"]=None
    else:
        a["mct_o"]=float(b[f"R{row}"].value)
        a["mct_w"]=format(a["mct_o"]*0.05, "0.3f")
    BOM.append(a)
#print(BOM)
outfile = open("BOM.json", "w")
json.dump(BOM, outfile, indent=3)
outfile.close()
wb.close




