import PySimpleGUI as sg
import shutil
import subprocess
import openpyxl
import pyautogui
import os

layout = [  [sg.Text('Inicar Bot')],
            [sg.Text('Arquivo Excel desejado: '),  sg.FileBrowse('Procurar',file_types=(("Excel Files", "*.xlsx"),))],
            [sg.Button('Iniciar'), sg.Button('Cancelar')] ]

window = sg.Window('Bot', layout)

def iniciarBot(planilha):
    
    planilhaExcel = openpyxl.load_workbook(planilha)['vendas']

    for linha in planilhaExcel.iter_rows(min_row=2):
        
        #Cliente
        pyautogui.click(677,424, duration=1)
        pyautogui.write(linha[0].value)

        #Produto
        pyautogui.click(679,449, duration=1)
        pyautogui.write(linha[1].value)

        #Quantidade
        pyautogui.click(697,479, duration=1)
        pyautogui.write(str(linha[2].value))

        #Categoria do Produto
        pyautogui.click(751,503, duration=1)
        pyautogui.write(linha[3].value)

        #salvarDiogo da Rocha
        pyautogui.click(636,526, duration=1)

        #OK
        pyautogui.click(717,486,duration=1)

#Fazer uma cópia do arquivo Excel nas pasta do Bot
def move_excel_file(file_path, destination_folder):
    try:
        shutil.copy(file_path, destination_folder)
        return 'OK'
    except Exception as e:
        return 'NO'

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancelar':
        break
    elif event == 'Iniciar':
        if values['Procurar']:
            caminho_planilha = values['Procurar'] #Obtém o caminho da planilha
            caminho_destino = 'X:\\XXX\\XXX' #Caminho destino da planilha
            if move_excel_file(caminho_planilha, caminho_destino) == 'OK':
                window.Minimize()
                subprocess.Popen(['pythonw', 'X:\\XXX\\XXX\\appSistema.pyw'])
                arquivoExcel = os.path.basename(caminho_planilha)
                caminho = os.path.join(caminho_destino, arquivoExcel)
                iniciarBot(caminho)

window.close()
