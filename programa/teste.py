import pyautogui as gui
import time
import cv2
import numpy as np
import datetime

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

def captura_tela():
    screenshot = gui.screenshot()
    imagem = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    return imagem

def redimensionar_template(imagem_template, largura, altura):
    return cv2.resize(imagem_template, (largura, altura))

def encontrar_template(imagem_tela, imagem_template, threshold=0.8):
    img_gray = cv2.cvtColor(imagem_tela, cv2.COLOR_BGR2GRAY)
    template_gray = cv2.cvtColor(imagem_template, cv2.COLOR_BGR2GRAY)
    
    resultado = cv2.matchTemplate(img_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    
    locais = np.where(resultado >= threshold)
    return zip(*locais[::-1])

# Acesso a notas fiscais de despesas
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

imagem_tela = captura_tela()
imagem_template = cv2.imread('C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/cliente_referencia.PNG')

# Obter a resolução da tela
resolucao_tela = gui.size()
largura_tela, altura_tela = resolucao_tela

# Redimensionar o template com base na resolução da tela
nova_largura = int(imagem_template.shape[1] * (1920 / largura_tela))
nova_altura = int(imagem_template.shape[0] * (1080 / altura_tela))  
imagem_template_redimensionada = redimensionar_template(imagem_template, nova_largura, nova_altura)

# Encontrar coordenadas do template redimensionado
coordenadas = encontrar_template(imagem_tela, imagem_template_redimensionada)

for (x, y) in coordenadas:
    gui.moveTo(x + (imagem_template_redimensionada.shape[1] // 0.8), y + (imagem_template_redimensionada.shape[0] // 2), duration=0.1)
    gui.click()
    gui.press("tab", presses=7)
    gui.press("enter")
    gui.write(cnpj)
    gui.press("enter", presses=2)
    break
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
gui.write("Colocar a descrição da nota")
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