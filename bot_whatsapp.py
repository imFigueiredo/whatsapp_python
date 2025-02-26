import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

workbook = openpyxl.load_workbook('clientes.xlsx')

pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

    #insira a menssagem que gostaria abaixo
    mensagem = f'Olá {nome}, tudo bem? 

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(30) 

    seta = pyautogui.locateCenterOnScreen('seta.png')
    if seta is not None:
        sleep(15) 
        pyautogui.click(seta[0], seta[1])
        sleep(15)
    else:
        print("Seta não encontrada na tela.")

    pyautogui.hotkey('ctrl', 'w')
    sleep(2) 
