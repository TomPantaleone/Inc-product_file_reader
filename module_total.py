#! Python 3.0
#! Inc-module_total.py
#! Module for file Inc-product_file_reader

productos= {
"ACELGA X PAQUETE      ":7,
"RADICHETA X ATADO      ":20,
"LECHUGA FRANCESA X KG.      ":6,
"RABANITO X ATADO.      ":15,
"LECHUGA MORADA X KG.      ":5,
"LECHUGA CRIOLLA x kg.      ":6,
"ESPINACA X ATADO      ":14,
"PEREJIL X ATADO      ":30,
"PUERRO X ATADO      ":20,
"albaca_modificar":25,
"BROCOLI x kg.      ":8,
"COLIFLOR x kg.      ":8,
"LECHUGA ESCAROLA X KG      ":5,
"HINOJO X KG.      ":10,
"LECHUGA CAPUCHINA X KG      ":8,
"LECHUGA MANTECOSA X KG.      ":5,
"CEBOLLA DE VERDEO X ATADO      ":30,
"APIO X KG.      ":10,
"REPOLLO BLANCO x kg.-      ":10,
"REPOLLO COLORADO X KG.      ":10,
"AKUSAY X KG      ":4,
"RUCULA X ATADO      ":30
}

def cant(prod_total_cell,cant_total_cell):
    if prod_total_cell in productos:
        print(str("\rcajones: ") + str(int(cant_total_cell / productos.get(prod_total_cell))))