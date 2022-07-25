#! Python 3.0
#! Inc-product_file_reader.py
#! File that reads a .xls and sorts info in differents lists.

from fileinput import close
from multiprocessing.sharedctypes import Value
import os, openpyxl
import pyexcel as p
import ast

pathfolder = (os.path.dirname(__file__))
filefolder = (pathfolder + "\\archivos\\")
workdir = os.chdir(filefolder)
fileinc = os.listdir()[0]
fileinc_ext = os.path.splitext(fileinc)

# Borrar archivos de la carpeta "pedidos"
for files in os.listdir(pathfolder + "\\pedidos\\"):
    os.remove(pathfolder + "\\pedidos\\" + files)

# Si el archivo es .xls, guardarlo como .xlsx y abrirlo. Si no lo es, abrirlo
if fileinc_ext[1] == ".xls":
    fileinc_ext2 = fileinc_ext[0] + "2.xlsx"
    p.save_book_as(file_name=fileinc, dest_file_name=fileinc_ext2)
    wb = openpyxl.load_workbook(fileinc_ext2)
    sheets = wb.sheetnames
    ws = wb[sheets[0]]
else:
    wb = openpyxl.load_workbook(fileinc)
    sheets = wb.sheetnames
    ws = wb[sheets[0]]

# Abre el xlsx de pedidos_base
wb_pedidos = openpyxl.load_workbook(pathfolder + "//listados//pedidos-base.xlsx")
ws_pedidos = wb_pedidos.active

    ### No Hardcodear los archivos a buscar. Resolverlo con una funcion
file = open(pathfolder + "\listados\productos_kilos.txt", "r")
contents = file.read()
productos = ast.literal_eval(contents)
file.close()

    ### Pensar si sirve cargar el listado de productos_kilos.txt y aplicarle un key-value de 0 a todo.
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

# Sacando datos para imprimir
suc_value = 0
for cell in range (2, ws.max_row-1):
    suc_cell = ws["A" + str(cell)].value                # Suc data
    oc_cell = (ws["B" + str(cell)].value)               # OC data
    prod_cell = (ws["H" + str(cell)].value)             # Producto data
    cant_cell = (ws["I" + str(cell)].value)             # Cantidad data

    # Chequea si la sucursal tiene productos en ifco o aca y los pone en distintas listas
    if suc_cell == None:
        continue
    elif suc_cell != suc_value:
        if suc_cell in sucursales_aca:
            list_prod = productos_aca
        else:
            list_prod = productos_ifco
        bultos = 0  

        # A la lista "Remito" le aplica como values las cantidades de los productos (Que pasaron en el listado) - Solo saca capuchina
        ws_pedidos["D1"] = oc_cell
        ws_pedidos["D2"] = suc_cell

        # Saca las cantidades finales (Cajones) - Solo
        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            bultos += prod_div
            list_prod[prod_cell] += prod_div
    # Saca cantidad (Archivo) de todos los productos
            for numero_pedido in range (1, ws_pedidos.max_row-1):
                prod_pedidos = ws_pedidos["C" + str(numero_pedido)].value
                if prod_pedidos == prod_cell:
                    ws_pedidos["B" + str(numero_pedido)] = cant_cell
                    ws_pedidos["F" + str(numero_pedido)] = prod_div

    else:
        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            bultos += prod_div
            list_prod[prod_cell] += prod_div
            for numero_pedido in range (1, ws_pedidos.max_row-1):
                prod_pedidos = ws_pedidos["C" + str(numero_pedido)].value
                if prod_pedidos == prod_cell:
                    ws_pedidos["B" + str(numero_pedido)] = cant_cell
                    ws_pedidos["F" + str(numero_pedido)] = prod_div

        if suc_cell != ws["A" + str(cell+1)].value:
            ws_pedidos["F31"] = bultos
            wb_pedidos.save(pathfolder + "//pedidos//sucursal_" + str(suc_cell) + ".xlsx")
            wb_pedidos = openpyxl.load_workbook(pathfolder + "//listados//pedidos-base.xlsx")
            ws_pedidos = wb_pedidos.active

    suc_value = suc_cell

wb_cantidad = openpyxl.load_workbook(pathfolder + "//listados//pedidos-base.xlsx")
ws_cantidad = wb_cantidad.active
for cantidad_pedidos in range (1, ws_cantidad.max_row-1):
    nomb_pedidos = ws_cantidad["C" + str(cantidad_pedidos)].value
    for x,y in productos_aca.items():
        if x == nomb_pedidos and y != 0:
            ws_cantidad["F" + str(cantidad_pedidos)] = y
    for x,y in productos_ifco.items():
        if x == nomb_pedidos and y != 0:
            ws_cantidad["G" + str(cantidad_pedidos)] = y
wb_cantidad.save(pathfolder + "//pedidos//cantidad.xlsx")
wb_cantidad.close()

file.close()
wb.close()