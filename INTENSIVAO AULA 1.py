#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[7]:


import pyautogui   #biblioteca de automação de comandos do mouse e teclado
import pyperclip   #biblioteca que vem junto com pyautogui; usada para digitar textos com caracteres especiais
import time   #biblioteca de funções relacionadas a tempo
import pandas as pd   #biblioteca para tratamento de arquivos
import openpyxl   #biblioteca para tratamento de arquivos excel; não ignora gráficos
import keyboard   #biblioteca de automação de comandos do mouse e teclado
import os
import shutil


# In[8]:


#2 - entrar no sistema da empresa (link do drive)

pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

pyautogui.hotkey('ctrl', 't')
pyperclip.copy(r'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

while not pyautogui.locateOnScreen('Logo Google Drive.PNG', grayscale=True, confidence=0.9):
    time.sleep(0.5)
time.sleep(1)

#3 - no sistema, encontrar a base de vendas (arquivo xlsx)

pyautogui.click(x=422, y=345, clicks=2, interval=0.20) 
time.sleep(2)
pyautogui.click(x=456, y=434) 
time.sleep(1)

#4 - download do arquivo

pyautogui.click(x=592, y=217)
time.sleep(4)

#5 - fechar a aba
pyautogui.hotkey('ctrl', 'w')


# In[13]:


# mover o arquivo para a pasta do projeto e deletá-lo
arquivos_downloads = os.listdir(r'C:\Users\W10\Downloads')

for arquivo in arquivos_downloads:
    if arquivo == 'Vendas - Dez.xlsx':
        local_arquivo = fr'C:\Users\W10\Downloads\{arquivo}'
        print(local_arquivo)
        local_final = os.getcwd()
        shutil.copy2(local_arquivo, local_final)
        
        time.sleep(2)
        
        os.remove(local_arquivo)


# In[19]:


#5 - importar para o py

df = pd.read_excel('Vendas - Dez.xlsx')   
display(df)
df.info()


# In[20]:


#6 - calcular os indicadores de interesse para a empresa: 'o faturamento e a quantidade de produtos vendidos no dia anterior'

# faturamento
dia_anterior = df['Data'].max()
df_dia_anterior = df.loc[df['Data']==dia_anterior, 'Valor Final'].to_frame()
faturamento = df_dia_anterior['Valor Final'].sum()
print(faturamento)

# quantidade de produtos
df_qtde_produtos = df.loc[df['Data']==dia_anterior, 'Produto'].to_frame()
qtde_produtos = len(df_qtde_produtos['Produto'].unique())
print(qtde_produtos)


# In[34]:


#7 - enviar um email para a diretoria com os indicadores de venda

#abrir email
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://outlook.live.com/mail/0/')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

#clicar em Novo Email
pyautogui.click(x=208, y=234)

#preencher as informações
pyautogui.write('bep_rafael@hotmail.com')   #como não tem caractere especial, sem problemas usar o pyautogui.write()
pyautogui.press('tab')   #como pyautogui.press('tab') não funcionou, recorri ao keyboard
keyboard.press_and_release('tab')
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
keyboard.press_and_release('tab')

texto = f'''
Prezada diretoria,

Segue relatório de vendas do mês:
Faturamento total do mês: R$ {faturamento:,.2f}.
Quantidade de produtos vendidos: {qtde_produtos:,} unidades.

Atenciosamente,
Rafael Muller.
'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)

#anexar arquivo no email
local_planilha = fr'{local_final}\Vendas - Dez.xlsx'

pyautogui.click(x=712, y=837)     #clicar em anexar
pyautogui.click(x=798, y=718)
time.sleep(2)
pyperclip.copy(local_planilha)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

#enviar email
pyautogui.hotkey('ctrl', 'enter')

