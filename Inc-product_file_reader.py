#! Python 3.0
#! Inc-product_file_reader.py
#! File that reads a .xls and sorts info in differents lists.

import os, openpyxl
import pyexcel as p
import ast

pathfolder = (os.path.dirname(__file__))
filefolder = (pathfolder + "\\archivos\\")
workdir = os.chdir(filefolder)
fileinc = os.listdir()[0]
fileinc_ext = os.path.splitext(fileinc)

if fileinc_ext[1] == ".xls":
    fileinc_ext2 = fileinc_ext[0] + "2.xlsx"
    p.save_book_as(file_name=fileinc, dest_file_name=fileinc_ext2)
    wb = openpyxl.load_workbook(fileinc_ext2)
    ws = wb.active              ###Cambiar para que elija la primer worksheet
else:
    wb = openpyxl.load_workbook(fileinc)
    ws = wb.active

 ### No Hardcodear por 3 los archivos a buscar. Resolverlo con una funcion
file = open(pathfolder + "\listados\productos_kilos.txt", "r")
contents = file.read()
productos = ast.literal_eval(contents)
file.close()

file = open(pathfolder + "\listados\productos_listado.txt", "r")
contents = file.read()
productos_ifco = ast.literal_eval(contents)
file.close()

file = open(pathfolder + "\listados\productos_listado.txt", "r")
contents = file.read()
productos_aca = ast.literal_eval(contents)
file.close()

file = open(pathfolder + "\listados\sucursales_aca.txt", "r")
contents = file.read()
sucursales_aca = ast.literal_eval(contents)
file.close()

file = open(pathfolder + "\listados\sucursales_ifco.txt", "r")
contents = file.read()
sucursales_ifco = ast.literal_eval(contents)
file.close()


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
        # print(prt_cell)
        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            # print(prod_div)
            productos_ifco[prod_cell] += prod_div
    else:
        prt_cell = f'''    Producto: {prod_cell} : {cant_cell}'''
        # print(prt_cell)     
        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            # print(prod_div)
            productos_ifco[prod_cell] += prod_div
    suc_value = suc_cell

print(productos_ifco)
print(productos_aca)

wb.close()