from pathlib import Path

import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Phone Number', size=(15,1)), sg.InputText(key='Phone Number')],
    [sg.Text('Email', size=(15,1)), sg.InputText(key='Email')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Combo Veg', size=(15,1)), sg.Combo(['Veg Manchurian Gravy with Classic Rice','Veg Manchurian Gravy with Classic Noodles','Chilly Paneer Gravy with Classic Rice','Chilly Paneer Gravy with Classic Noodles'], key='Combo Veg')],
    [sg.Text('Transaction', size=(15,1)),
                            sg.Checkbox('Phone Pe', key='Phone Pe'),
                            sg.Checkbox('GooglePay', key='GooglePay'),
                            sg.Checkbox('Net Banking', key='Net Banking')],
    [sg.Text('No. of Orders', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Orders')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Data Entry Form', layout)

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
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()