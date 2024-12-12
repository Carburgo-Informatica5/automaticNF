# importando bibliotecas 
import pyautogui as gui
import sqlite3

cnpj_dest = "91667618000165"

# Conexão com o banco
conn = sqlite3.connect('banco de dados.db')
cursor = conn.cursor()

# Consulta no banco de dados
query = "SELECT empresa, revenda FROM ger_revenda WHERE cnpj=?"
cursor.execute(query, (cnpj_dest))
result = cursor.fetchone()

if result:
    empresa, revenda = result
    print(f"Empresa: {empresa}, Revenda: {revenda}")
    
    gui.alert("Teste de código")
    
    # Acessando o menu
    gui.press("alt")
    gui.press("right")
    gui.press("down")
    gui.press("enter")
    gui.press("down", presses=2)
    
    gui.press(str(empresa), presses=int(revenda))
else:
    print("Erro, CNPJ não encontrado")