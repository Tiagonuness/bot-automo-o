import sys
import openpyxl
import pyautogui
import keyboard
    
planilha = openpyxl.load_workbook('Projeto1/vendas_de_produtos.xlsx')['vendas']

def check_escape_key():
    return keyboard.is_pressed('esc')

if len(sys.argv) != 2:
    print("Por favor, forneça o caminho do arquivo Excel como argumento.")
    sys.exit(1)

excel_file_path = sys.argv[1]


for linha in planilha.iter_rows(min_row=2):

    if check_escape_key():
        print("Tecla 'ESC' pressionada. Parando a execução.")
        break
    
    #Cliente
    pyautogui.click(677,424, duration=1)
    cliente = pyautogui.write(linha[0].value)

    #Produto
    pyautogui.click(679,449, duration=1)
    produto = pyautogui.write(linha[1].value)

    #Quantidade
    pyautogui.click(697,479, duration=1)
    quantidade = pyautogui.write(str(linha[2].value))

    #Categoria do Produto
    pyautogui.click(751,503, duration=1)
    categoria_do_produto = pyautogui.write(linha[3].value)

    #salvarDiogo da Rocha
    pyautogui.click(636,526, duration=1)

    #OK
    pyautogui.click(717,486,duration=1)