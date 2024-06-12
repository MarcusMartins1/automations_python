# Importa as bibliotecas necessárias
import pandas as pd
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Abre o WhatsApp Web no navegador padrão
webbrowser.open('https://web.whatsapp.com/')
# Aguarda 10 segundos para garantir que a página carregue completamente
sleep(10)

# Carrega a planilha Excel que contém os contatos --- pode ser alterado para sua planilha
workbook = openpyxl.load_workbook('contatos.xlsx')

# Seleciona a primeira planilha do arquivo Excel
pagina_clientes = workbook['Planilha1']

# Loop através de cada linha da planilha, começando da segunda linha
for linha in pagina_clientes.iter_rows(min_row=2):

    # Extrai o nome, número e mensagem de cada linha da planilha
    nome = linha[0].value
    numero = linha[1].value
    mensagem = linha[2].value

    # Formata a mensagem a ser enviada
    msg = f'Opa, {nome}. {mensagem}'

    # Cria o link para enviar a mensagem pelo WhatsApp Web
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={numero}&text={quote(msg)}'

    # Abre o link no navegador
    webbrowser.open(link_mensagem_whatsapp)
    
    # Aguarda 15 segundos para garantir que a página do WhatsApp Web carregue completamente
    sleep(15)
    
    try:
        # Localiza o ícone da seta para enviar a mensagem na tela --- necessário criar um arquivo com o print da seta
        seta = pyautogui.locateCenterOnScreen('seta.png')
        # Aguarda 8 segundos para garantir que a localização da seta seja identificada
        sleep(8)
        # Clica na seta para enviar a mensagem
        pyautogui.click(seta[0], seta[1])
        # Aguarda 8 segundos para garantir que a mensagem seja enviada
        sleep(8)
        # Fecha a aba do navegador
        pyautogui.hotkey('ctrl', 'w')
        # Aguarda 8 segundos para garantir que a aba seja fechada
        sleep(8)
    except:
        # Caso não seja possível enviar a mensagem, imprime um aviso no console
        print(f'Não foi possível enviar mensagem para {nome} do número {numero}')
        # Registra o erro em um arquivo CSV
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{numero}\n')
