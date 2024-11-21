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


gui.alert("O código vai começar. Não utilize nada do computador até o código finalizar!")


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
gui.press("tab")



# posicao_avancado = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/avancado.jpg", confidence=0.8)
# gui.click(posicao_avancado, duration=0.1)
# gui.write(cnpj_eminente)
# gui.press(["enter", "enter", "tab", "tab"])

# posicao_entrada= gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/data_entrada.jpg", confidence=0.8)
# gui.click(posicao_entrada, clicks=2, duration=0.1)
# gui.press("tab")
# gui.write(data_emi)

# posicao_modelo = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/modelo.jpg", confidence=0.8)
# gui.click(posicao_modelo, duration=0.1)
# gui.write(modelo)
# gui.press("tab")

# posicao_chave = gui.locateCenterOnScreen("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/programa/assets/chaveAcesso.jpg", confidence=0.8)
# gui.click(posicao_chave, duration=0.1)
# gui.write(chave_acesso)
# gui.press("enter")