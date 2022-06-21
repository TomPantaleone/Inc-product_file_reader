#! Python 3.0
#! Inc-product_file_reader.py
#! File that reads a .xls and sorts info in differents lists.

import os, openpyxl
import pyexcel as p
import module_total

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

def cant(prod_total_cell,cant_total_cell):
    if prod_total_cell in produc:
        print(int(cant_total_cell / produc.get(prod_total_cell)))

# Diccionario para cantidades sumadas
productos_cant = {}

# Sacando datos para imprimir
suc_value = 0
for cell in range (2, ws.max_row-1):
    suc_cell = ws["A" + str(cell)].value                # Suc data
    oc_cell = (ws["B" + str(cell)].value)               # OC data
    prod_cell = (ws["H" + str(cell)].value)             # Producto data
    cant_cell = (ws["I" + str(cell)].value)             # Cantidad data
    if suc_cell == None:
        continue
    elif suc_cell != suc_value:
        prt_cell = f'''Sucursal: {suc_cell}
  OC: {str(oc_cell)}
    Producto: {prod_cell} : {cant_cell}'''
        print(prt_cell)
        if module_total.cant(prod_cell,cant_cell) == True:
            print(module_total.cant(prod_cell,cant_cell))
    else:
        prt_cell = f'''    Producto: {prod_cell} : {cant_cell}'''
        print(prt_cell)     
        # if module_total.cant(prod_cell,cant_cell) == True:
        #     module_total.cant(prod_cell,cant_cell)
    suc_value = suc_cell

print(productos_cant)
wb.close()