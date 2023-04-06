#Importar librería
import pandas as pd
import PySimpleGUI as sg

#Asignar tema(apariencia)
sg.theme('DarkTeal9')

#Agregar ruta del archivo excel
excel_file = 'data_entry_form.xlsx'
df = pd.read_excel(excel_file)


#Crear layout
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    #Combo box
    [sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue','Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15,1)),
    sg.Checkbox('German', key='German'),
    sg.Checkbox('Spanish', key='Spanish'),
    sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                        initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]









#Crear Ventana
window = sg.Window('Simple data entry form', layout)

#Definir funcionalidad boton limpiar
def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(excel_file, index=False)
        sg.popup('Data_saved!')
        clear_input()
window.close()