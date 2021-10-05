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


# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# Para instalar as bibliotecas, use o 

#get_ipython().system('pip install pyautogui')

#get_ipython().system('pip install pyperclip')

# Caso você não esteja no jupyter, precisará instalar as dependências/bibliotecas: pandas, numpy, openpyxl

import pandas as pd
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no link)

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar até o local do relatório (pasta exportar)

pyautogui.click(x=304, y=341, clicks=2)

time.sleep(2)

# Passo 3: Fazer o download do relatório

pyautogui.click(x=304, y=331, button='right')
pyautogui.click(x=460, y=652)

time.sleep(3)

# Passo 4: Calcular os indicadores

tabela = pd.read_excel(r"C:/Users/CARECA/Downloads/Vendas - Dez.xlsx")

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()


# Passo 5: Entrar no email

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 6: Enviar por email o resultado

pyautogui.click(x=38, y=166)
time.sleep(9)

pyautogui.write("alison002h@gmail.com")

pyautogui.press("tab") #seleciona o email;
pyautogui.press("tab") #pula para a seção "assunto";

pyperclip.copy("Relatório Semanal (Exemplo aula 1 - Python)")

pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") #pula para a seção do corpo do email;

texto = f"""
Bom dia, prezados

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}
    
Abs
Alison Chs """

pyperclip.copy(texto)

pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl", "enter")

