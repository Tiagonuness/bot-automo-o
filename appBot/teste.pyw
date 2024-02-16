import PySimpleGUI as sg
import shutil
import subprocess
import openpyxl
import pyautogui
import keyboard
import os

sg.theme('Dark')

layout = [  [sg.Text('Inicar Bot')],
            [sg.Text('Arquivo excel desejado: '),  sg.FileBrowse('Procurar',file_types=(("Excel Files", "*.xlsx"),))],
            [sg.Button('Iniciar'), sg.Button('Cancelar')] ]

# Create the Window
window = sg.Window('Bot', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancelar': # if user closes window or clicks cancel
        break
    elif event == 'Iniciar':
        pass



window.close()