#! Python 3.0
#! Inc-product_file_reader.py
#! File that reads a .xls and sorts the info in differents lists.

import os, openpyxl
import pyexcel as p

pathfolder = (os.path.dirname(__file__))
filefolder = (pathfolder + "\\archivos\\")
workdir = os.chdir(filefolder)
fileinc = os.listdir()[0]
fileinc_ext = os.path.splitext(fileinc)

if fileinc_ext[1] == ".xls":
    fileinc_ext2 = fileinc_ext[0] + "2.xlsx"
    p.save_book_as(file_name=fileinc, dest_file_name=fileinc_ext2)
    wb = openpyxl.load_workbook(fileinc_ext2)
    ws = wb.active #Cambiar para que elija el primer worksheet
else:
    wb = openpyxl.load_workbook(fileinc)
    ws = wb.active

# Sacando SUCURSAL
suc_value = 0
for suc_row in ws.iter_rows(min_row=2, max_row=ws.max_row -1, max_col=1):
    for cell in suc_row:
        if cell.value != suc_value:
            print(cell.value)
            
            
            # Hacer funcion
            # for oc_row in ws.iter_rows(min_row=2, max_row=ws.max_row -1, max_col=2):
            #     for cell in oc_row:
            #         print(cell.value)
            #         continue

        suc_value = cell.value


wb.close()