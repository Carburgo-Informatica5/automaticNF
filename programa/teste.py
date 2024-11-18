import pyautogui as gui
import time

# Substituir pelas variaveis da nota
nmr_nota = "39216"
serie = "1"
contador = "0"
cnpj_eminente = "15000312000172"
chave_acesso = "43241115000312000172550010000392161402085460"
data_emi = "04112024"
modelo = "55"

gui.PAUSE = 0.5

time.sleep(5)
gui.press("alt")
gui.press("right", presses=6)
gui.press("down", presses=4)
gui.press("enter")
gui.press("down", presses=2)
gui.press("enter")
time.sleep(5)

posicao_transacao = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/transacao.jpg", confidence=0.8)
posicao_nmr_nota = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/nmr_nota.jpg", confidence=0.8)
posicao_serie = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/serie.jpg", confidence=0.8)
posicao_contador = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/contador.jpg", confidence=0.8)

gui.click(posicao_transacao, duration=0.1)
gui.press("down")
gui.press("enter")
gui.click(posicao_nmr_nota, duration=0.1)
gui.write(nmr_nota)
gui.press("enter")
gui.click(posicao_serie, duration=0.1)
gui.write(serie)
gui.press("enter")
gui.click(posicao_contador, duration=0.1)
gui.write(contador)
gui.press("enter")

posicao_mao = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/mao.jpg", confidence=0.8)
gui.click(posicao_mao, duration=0.1)

posicao_avancado = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/avancado.jpg", confidence=0.8)
gui.click(posicao_avancado, duration=0.1)
gui.write(cnpj_eminente)
gui.press(["enter", "enter", "tab", "tab"])

posicao_entrada= gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/data_entrada.jpg", confidence=0.8)
gui.click(posicao_entrada, clicks=2, duration=0.1)
gui.press("tab")
gui.write(data_emi)

posicao_modelo = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/modelo.jpg", confidence=0.8)
gui.click(posicao_modelo, duration=0.1)
gui.write(modelo)
gui.press("tab")

posicao_chave = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/chaveAcesso.jpg", confidence=0.8)
gui.click(posicao_chave, duration=0.1)
gui.write(chave_acesso)
gui.press("enter")