import pandas as pd
import PySimpleGUI as sg

sg.theme('DarkPurple6')

EXCEL_FILE = "C:/Users/User/Documents/Buku Tamu/IdentitasTamu.xlsx"; #Change With Your Directory Excel File

df = pd.read_excel(EXCEL_FILE)

layout=[
    [sg.Text('Masukan Identitas Anda: ')],
    [sg.Text('Nama: ', size=(15,1)), sg.InputText(key='Nama')],
    [sg.Text('No.Telp: ', size=(15,1)), sg.InputText(key='Telp')],
    [sg.Text('Alamat: ', size=(15,1)), sg.Multiline(key='Alamat')],
    [sg.Text('Jenis Kelamin: ', size=(15,1)), sg.Combo(['pria', 'wanita'],key='kelamin')],
    [sg.Text('Keterangan: ', size=(15,1)), sg.Checkbox('Hadir',key='hadir'),
                                            sg.Checkbox('Tidak Hadir',key='tidakhadir'),],
    [sg.Submit(), sg.Button('clear'), sg.Exit()]
]

window=sg.Window('Buku Tamu', layout)

def clear_input():
    for key in values:
        window[key]('')
        return None
    
while True :
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Identitas Berhasil Di Simpan')
        clear_input()
window.close()
