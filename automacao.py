import pyautogui as gui
import pyperclip
from time import sleep
import pandas as pd
import numpy as np


gui.PAUSE = 1

gui.press('win')
gui.write('chrome')
gui.press('enter')
gui.alert("O programa irá começar. Aperte Ok quando estiver pronto!")
gui.hotkey('ctrl', 't')

link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing' #Inserir banco de dados
pyperclip.copy(link)
gui.hotkey('ctrl', 'v')
gui.press('enter')

sleep(5)

gui.click(x=351, y=304, clicks=2)
sleep(2)
gui.click(x=372, y=477)
sleep(2)
gui.click(x=1159, y=192)
sleep(2)
gui.click(x=941, y=623)
sleep(5)
 
######################################

dataframe = pd.read_excel(r'Vendas - Dez.xlsx') #Endeço do arquivo baixado
print(dataframe)

faturamento = dataframe['Valor Final'].sum()
qtde_produtos = dataframe['Quantidade'].sum()
#print(faturamento)
#print(qtde_produtos)

####################################################

gui.hotkey('ctrl', 't')
gui.write('https://mail.google.com/') #Abre o Gmail - Importante está logado na máquina
gui.press('enter')
sleep(5)

gui.click(x=104, y=209)
sleep(2)

gui.write('') #Inserir o Email de destino
gui.press('tab')
gui.press('tab')
sleep(2)
assunto = 'Relatório de vendas de ontem' #Inserir o assunto do email
pyperclip.copy(assunto)
gui.hotkey('ctrl', 'v')
gui.press('tab')
sleep(2)

texto = f'''Bom dia, pessoal!
Esse é um email teste.

O faturamento de ontem foi de: R${faturamento:,.2f}
E a quatidade de vendas de ontem foi de: {qtde_produtos:,}

Abs,
Rayan.'''

pyperclip.copy(texto)
gui.hotkey('ctrl', 'v')
gui.hotkey('ctrl', 'enter') #Atalho de teclado para enviar o email no GMAIL


