import pyautogui as gui
import time
import datetime
import pygetwindow as gw
import pytesseract
from PIL import Image


# Substituir pelas variaveis da nota
nmr_nota = "3921"
serie = "1"
contador = "0"
nome_eminente = "MEGA COMERCIAL DE PRODUTOS ELETRONICOS LTDA"
cnpj = "15000312000172"
chave_acesso = "43241115000312000172550010000392161402085460"
data_emi = "04112024"
modelo = "55"
valor_total = "2000"
data_vali = "24122024"

gui.PAUSE = 0.5
data_atual = datetime.datetime.now()

data_formatada = data_atual.strftime("%d%m%Y")

gui.alert("O código vai começar. Não utilize nada do computador até o código finalizar!")

time.sleep(5)

# Localiza a janela do BRAVOS pelo título
window = gw.getWindowsWithTitle('BRAVOS v5.17 Evolutivo')[0]  # Assumindo que é a única com "BRAVOS" no título

# Centraliza a janela se necessário
window.activate()

# Calcula a posição relativa do ícone na barra de ferramentas
x, y = window.left + 275, window.top + 80  # Ajuste os offsets conforme necessário
time.sleep(3)
gui.moveTo(x, y, duration=0.5)
gui.click()
time.sleep(4)
gui.press("tab", presses=19)
gui.write(cnpj)
time.sleep(2)
gui.press("enter")

pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR/tesseract.exe'

nova_janela = gw.getActiveWindow()

janela_left = nova_janela.left
janela_top = nova_janela.top

time.sleep(5)

x, y, width, height = janela_left + 500, janela_top + 323, 120, 21  

screenshot = gui.screenshot(region=(x, y, width, height))

screenshot = screenshot.convert("L")  
threshold = 150  
screenshot = screenshot.point(lambda p: p > threshold and 255)  

config = r'--psm 7 outputbase digits'

cliente = pytesseract.image_to_string(screenshot, config=config)

time.sleep(5)
gui.press("alt")
gui.press("right", presses=6)
gui.press("down", presses=4)
gui.press("enter")
gui.press("down", presses=2)
gui.press("enter")
time.sleep(5)

# Preenche campos de nmr da nota, série, transação, contador e cliente
gui.press("tab")
gui.write(nmr_nota)
gui.press("tab")
gui.write(serie)
gui.press("tab")
gui.press("down")
gui.press("tab")
gui.write(contador)
gui.press("tab")
gui.write(cliente)
gui.press("tab")
gui.press("enter")

gui.press("tab", presses=5)
gui.write(data_formatada)
gui.press("tab")
gui.write(data_emi)
gui.press("tab", presses=5)
gui.write("9")
gui.press("tab", presses=2)
gui.write("5148")
gui.press("tab", presses=19)
gui.write(chave_acesso)
gui.press("tab")
gui.write(modelo)
gui.press("tab", presses=18)
gui.press("right", presses=2)
gui.press("tab", presses=5)
gui.press("enter")
gui.press("tab", presses=4)
gui.write("7")
gui.press("tab", presses=10)
gui.write("1")
gui.press("tab")
gui.write(valor_total)
gui.press("tab", presses=26)
gui.write("TESTE DANI LANÇAMENTO DE NOTA") # Variavel Relativa puxar pelo email
gui.press(["tab", "enter"])
gui.press("tab", presses=11)
gui.press("left", presses=2)
gui.press("tab", presses=5)
gui.press(["enter", "tab", "enter"])
gui.press("tab", presses=5)
gui.write(data_vali)
gui.press("tab", presses=4)
gui.press(["enter", "tab", "tab", "tab", "enter"])
gui.press("tab", presses=36)
gui.press("enter")
gui.moveTo(1150, 758)
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.click()
gui.press("enter")
gui.press("tab", presses=8)
gui.write("1") # Aqui várialvel relativa Revenda centro de custo puxar email
gui.press("tab", presses=3)
gui.write("1") # Aqui várialvel relativa de centro de custo puxar email
gui.press("tab", presses=2)
gui.write("5148") # Aqui várialvel relativa de Origem puxar email
gui.press("tab", presses=3)
gui.write("100") # Aqui várialvel relativa de Percentual puxar email
gui.press("tab")
gui.moveTo(1202, 758)
gui.click()
gui.moveTo(1271, 758)
gui.click()
gui.press("tab", presses=3)
# gui.press("enter") Se tirar do comentario ele grava a nota no sistema