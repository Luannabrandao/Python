# Importar as bibliotecas
import pyautogui # faz a automação do mouse e teclado
import time # controla o tempo do nosso porgrama 
import pyperclip #ela permite a gente copiar e colar com python

pyautogui.PAUSE = 1

pyautogui.alert("O Programa vai começar NÃO USE seu computador enquanto isso")

# Passo 1: Entrar no sistema (link do google Drive)
pyautogui.hotkey('ctrl', 't')
link = ('https://drive.google.com/drive/shared-with-me')
# da crtl+c e crtl+v
pyperclip.copy(link)
#cola o link
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
# esperar um pouco
time.sleep(5)
# Passo 2 : Entrar na pasta da Aula 1 
pyautogui.click(351, 335, clicks= 2)
# Passo 3: Fazer o download da base de vendas
pyautogui.click(1028, 372)
pyautogui.click(1160, 192)
pyautogui.click(1075, 555)
time.sleep(10)
time.sleep(5)
pyautogui.position()
# Passo 4: calcular os indicadores (faturamento e a Quantidade de produtos)
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Clara\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
qtde_produtos= tabela["Quantidade"].sum()

# Passo 5: entrar no meu email
# email da diretoria: luannav.s.santos+diretoria@gmail.com
pyautogui.hotkey('ctrl', 't')
link = "https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox"
# da crtl+c e crtl+v
pyperclip.copy(link)
#cola o link
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
# esperar um pouco
time.sleep(8)

# Passo 6: criar email
pyautogui.click(108, 499)
time.sleep(5)
# escolher o e-mail
pyautogui.write("luannav.s.santos+diretoria@gmail.com")
pyautogui.press("tab")
#Preencher o assunto
pyautogui.press("tab")
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")

# preencher o email
pyautogui.press("tab")
texto_email = f""" 
Prezados, bom dia 

O faturameto foi de: R$:{faturamento}
A quantidade de produtos foi de: {qtde_produtos}

Abs 
Luanna
"""
pyperclip.copy(texto_email)
pyautogui.hotkey("ctrl", "v")
# Passo 7:enviar e-mail
pyautogui.hotkeu("ctrl", "enter")