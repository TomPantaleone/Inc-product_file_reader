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


# archivos_text("\listados\productos_listado.txt", "productos_aca")

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

file = open(pathfolder + "/listados//remito.txt", "r")
contents = file.read()
remito = ast.literal_eval(contents)

# Funciones
def dict_cant(num_y,cant_var):                                         # Para poner totales de productos en los diccionarios (cajones y bultos)
    for x,y in remito.items():
        if x == prod_cell:
            y[num_y] = cant_var

def archivo_txt(nombre_archivo,listado):                        # Para crear .txt de los diccionarios
    file_txt = open(pathfolder + nombre_archivo, "w")
    for x,y in listado.items():
        if y != [0,0]:
            file_txt.write((x) + ":" + str(y) + "\n")
    file_txt.close

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
        remito["ORDEN"] = oc_cell
        remito["SUCURSAL"] = suc_cell       
        dict_cant(0,cant_cell)

        # Saca las cantidades finales (Cajones) - Solo
        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            bultos += prod_div
            list_prod[prod_cell] += prod_div
        dict_cant(1,prod_div)
    # Saca cantidad (Archivo) de todos los productos
    else:
        dict_cant(0,cant_cell)

        if prod_cell in productos:
            prod_div = int((cant_cell / productos.get(prod_cell)))
            bultos += prod_div
            list_prod[prod_cell] += prod_div
        dict_cant(1,prod_div)

        if suc_cell != ws["A" + str(cell+1)].value:
            remito["BULTOS"] = bultos
            archivo_txt(("/pedidos//sucursal_" + str(suc_cell) + ".txt"),remito)
            remito = ast.literal_eval(contents)


    suc_value = suc_cell

archivo_txt("/pedidos//cantidad_aca.txt",productos_aca)
archivo_txt("/pedidos//cantidad_ifco.txt",productos_ifco)
file.close()
wb.close()