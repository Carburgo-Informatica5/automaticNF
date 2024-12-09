# Seleção de revendo por meio do cpf
import pyautogui as gui
cnpj_dest = "91667618000750"

gui.alert("código de teste")

"""Menu de alteração da revenda"""
gui.press("alt")
gui.press("right")
gui.press("down")
gui.press("enter")
gui.press("down", presses=2)

"""Atrelar o CNPJ de cada filial para cada variável e fazer isso de acordo com a nota"""

match cnpj_dest:
    case "08680053000102":
        gui.press("1") # Revenda 1.1
    case "91667618000165":
        gui.press("2") # Revenda 2.1
    case "91667618000599":
        gui.press("2", presses=2) # Revenda 2.2
    case "91667618000670":
        gui.press("2", presses=3) # Revenda 2.3
    case "91667618000750":
        gui.press("2", presses=4) # Revenda 2.4
    case "91667618000831":
        gui.press("2", presses=5) # Revenda 2.5
    case "02918557000131":
        gui.press("3") # Revenda 3.1
    case "02918557000212":
        gui.press("3", presses=2) # Revenda 3.2
    case "02918557000301":
        gui.press("3", presses=3) # Revenda 3.3
    case "09112414000187":
        gui.press("4") # Revenda 4.1
    case "09112414000268":
        gui.press("4", presses=2) # Revenda 4.2
