import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


webbrowser.open('https://web.whatsapp.com/')
sleep(10)
workbook = openpyxl.load_workbook("clientes.xlsx")
pagina_clientes = workbook['Planilha1']


for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome}, seu boleto vence {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.tecsigma.com.br'
    link_msg_wtsp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_msg_wtsp)
    try:
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(2)
        pyautogui.click(seta[0], seta[1])
        sleep(2)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    except:
        print(f'Não foi possível enviar a mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
