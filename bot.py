import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep


# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
clientes = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = clientes['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    try:
        sleep(1)
        nome = linha[0].value.capitalize()
        telefone = linha[1].value
        vencimento = linha[2].value.strftime('%d/%m/%Y')
        mensagem = f'Olá {nome} seu boleto vence no dia {vencimento}.'

        # Criar links personalizafdos do whatsapp e enviar mendagens para cada cliente com base nos dados da planilha
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={mensagem}'

        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.press('enter')
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
    except:
        print(f'Não foi possível enviar mensagem para {nome}/n')
        with open('erro.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
