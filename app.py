import openpyxl
import webbrowser
import pyautogui
from urllib.parse import quote
from time import sleep
import os 

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes =  workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_do_pagamento.com'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente com base nos dados da planilha

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        seta = pyautogui.locateCenterOnScreen('seta1.png')
        sleep(20)
        pyautogui.click(seta[0],seta[1])
        sleep(20)
        pyautogui.hotkey('ctrl','w')
        sleep(10)

        # Salvar contatos em que não foi possível enviar mensagem

    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')