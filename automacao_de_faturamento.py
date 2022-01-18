# # biblioteca necessaria para fazer download, a fim de automatizar os processos:
# #
# # pyautogui: Para automatizar com comandos do teclado e do mouse.
# # pyperclip: para auxiliar o pyautogui.https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga

from email import enviar_email
import pandas as pd
from time import sleep
import pyautogui as auto  # opcional (from pyautogui import press)
import pyperclip
# Desafio:
# Com o chrome fechado:
auto.PAUSE = 1  # Código para fazer a pausa de um comando para o outro com tempo em segundo
auto.press("win")
auto.write("chrome")
auto.press("enter")
# Código que deve ser usado em um editor web.
# =>Entrar no site da empresa:
auto.hotkey("ctrl", "t")  # para usar um atalho no browser
url = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
# para tratar o caso de caracteres especiais e copiando a url
pyperclip.copy(url)
auto.hotkey("ctrl", "v")
# auto.write(url)  # para escrever a url desejada no navegador
auto.press("enter")  # para pressionar alguma tecla

sleep(2)  # Para fazer o sistema esperar carregar por 5 segundos
auto.click(x=398, y=279, clicks=2)

auto.click(x=398, y=279)
auto.click(x=1065, y=163)
sleep(5)
auto.click(x=905, y=588)

# Importando e trabalhando com base de dados

tabela_base = pd.read_excel(r"C:\Users\dafs\Downloads\Vendas - Dez.xlsx")
# O "r" serve para ler exatamente o que foi escrito, sem ler os caracteres especiais.
print(tabela_base)
faturamento = tabela_base['Valor Final'].sum()
quantidade_produtos = tabela_base['Quantidade'].sum()

print(faturamento)

enviar_email([faturamento, quantidade_produtos])
