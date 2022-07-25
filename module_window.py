import PySimpleGUI as sg      
import os
import ast

pathfolder = (os.path.dirname(__file__))

sucursal = {
"ACELGA X PAQUETE      ":1,
"RADICHETA X ATADO      ":2,
"LECHUGA FRANCESA X KG.      ":3,
"RABANITO X ATADO.      ":4,
"LECHUGA MORADA X KG.      ":5,
"LECHUGA CRIOLLA x kg.      ":6,
"ESPINACA X ATADO      ":7,
"PEREJIL X ATADO      ":8,
"PUERRO X ATADO      ":9,
"REMOLACHA X KG.      ":10,
"ALBAHACA X ATADO      ":11,
"COLIFLOR x kg.      ":12,
"BROCOLI x kg.      ":13,
"LECHUGA ESCAROLA X KG      ":14,
"HINOJO X KG.      ":16,
"LECHUGA CAPUCHINA X KG      ":16,
"LECHUGA MANTECOSA X KG.      ":17,
"CEBOLLA DE VERDEO X ATADO      ":18,
"APIO X KG.      ":19,
"AKUSAY X KG      ":20,
"REPOLLO COLORADO X KG.      ":21,
"REPOLLO BLANCO x kg.-      ":22,
"RUCULA X ATADO      ":23
}

# print("type producto: ")
# prod = input()

# for x,y in sucursal.items():
#     if x == prod:
#         continue
#     print(x + " : " + str(y))

### ---- window screen
layout = [[sg.Text('Elejir que productos no se van a contabilizar')],      
            [sg.Checkbox("Acelga", key="ACELGA X PAQUETE      ")],      
            [sg.Checkbox("Radicheta", key="RADICHETA X ATADO      ")],   
            [sg.Checkbox("Lechuga Francesa", key="LECHUGA FRANCESA X KG.      ")],
            [sg.Checkbox("Rabanito", key="RABANITO X ATADO.      ")],
            [sg.Checkbox("Lechuga morada", key="LECHUGA MORADA X KG.      ")],
            [sg.Checkbox("Lechuga criolla", key="LECHUGA CRIOLLA x kg.      ")],
            [sg.Checkbox("Espinaca", key="ESPINACA X ATADO      ")],
            [sg.Checkbox("Perejil", key="PEREJIL X ATADO      ")],
            [sg.Checkbox("Puerro", key="PUERRO X ATADO      ")],   
            [sg.Checkbox("Remolacha", key="REMOLACHA X KG.      ")],
            [sg.Checkbox("Albahaca", key="ALBAHACA X ATADO      ")],
            [sg.Checkbox("Coliflor", key="COLIFLOR x kg.      ")],
            [sg.Checkbox("Brocoli", key="BROCOLI x kg.      ")],
            [sg.Checkbox("Escarola ancha", key="LECHUGA ESCAROLA X KG      ")],
            [sg.Checkbox("Hinojo", key="HINOJO X KG.      ")],
            [sg.Checkbox("Lechuga Capuchina", key="LECHUGA CAPUCHINA X KG      ")],
            [sg.Checkbox("Lechuga Mantecosa", key="LECHUGA MANTECOSA X KG.      ")],
            [sg.Checkbox("Cebolla de verdeo", key="CEBOLLA DE VERDEO X ATADO      ")],
            [sg.Checkbox("Apio", key="APIO X KG.      ")],
            [sg.Checkbox("Akusay", key="AKUSAY X KG      ")],
            [sg.Checkbox("Repollo colorado", key="REPOLLO COLORADO X KG.      ")],
            [sg.Checkbox("Repollo blanco", key="REPOLLO BLANCO x kg.-      ")],
            [sg.Checkbox("Rucula", key="RUCULA X ATADO      ")],
            [sg.Submit("Aceptar"), sg.Cancel("Salir")]]      

window = sg.Window('automatizador de pedidos', layout)    

while True:
    event, values = window.read() 
    print(values)
    if event == "Aceptar":
        for x,y in values.items():
            if y == True:
                if x in sucursal.keys():
                    sucursal.pop(x)
        sg.popup('You entered', sucursal)
        break
    elif event == sg.WIN_CLOSED or event == "Salir":
        break

window.close()

