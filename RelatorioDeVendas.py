!pip install pandas
!pip install pyautogui
!pip install pyperclip

import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=329, y=289, clicks=2)
time.sleep(2)
pyautogui.click(x=329, y=289)
pyautogui.click(x=1154, y=188)
pyautogui.click(x=1133, y=630)
time.sleep(8)

tabela = pd.read_excel(r"C:\Users\YourUser\Downloads\Vendas - Dez.xlsx")
display(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

link = ("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "t")
pyautogui.write(link)
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=91, y=210)
time.sleep(0.5)

pyautogui.write("YourEmail@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
time.sleep(0.5)
pyautogui.write("Relatório de vendas.")
pyautogui.press("tab")

texto = f"""Prezado Mr.
Se você recebeu este email o processo foi concluido com sucesso
O faturamento total em R$ foi de: {faturamento}
A quantidade de produtos vendidos foi de: {quantidade}

Parabéns, você conseguiu."""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl","enter")
