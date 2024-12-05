import pyautogui as gui
import time
import datetime
import pygetwindow as gw
import pytesseract
from email import parser
from email.header import decode_header
import os
import poplib

# Configurações do servidor e credenciais
HOST = 'mail.carburgo.com.br'  # Servidor POP3 corporativo
PORT = 995                    # Porta segura (SSL)
USERNAME = 'caetano.apollo@carburgo.com.br'
PASSWORD = 'p@r!sA1856'

# Assunto alvo para busca
ASSUNTO_ALVO = 'Lançamentos notas fiscais DANI'

# Diretório para salvar os anexos
DIRECTORY = 'anexos'  # Pasta onde os anexos serão salvos

# Funções auxiliares

def decode_header_value(header_value):
    """Função auxiliar para decodificar o header do e-mail."""
    decoded, encoding = decode_header(header_value)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else 'utf-8', errors='ignore')
    return decoded

def decode_body(payload, charset):
    """Função auxiliar para decodificar o corpo do e-mail."""
    if charset is None:
        charset = 'utf-8'
    try:
        return payload.decode(charset)
    except (UnicodeDecodeError, LookupError):
        return payload.decode('ISO-8859-1', errors='ignore')

def extract_values(text):
    """Extrai os valores do corpo do e-mail."""
    values = {
        'departamento': None,
        'origem': None,
        'descricao': None,
        'revenda_cc': None,
        'cc': None,
        'rateio': None,
        'cod_item': None
    }

    lines = text.splitlines()
    for line in lines:
        if line.startswith("departamento:"):
            values['departamento'] = line.split(":", 1)[1].strip()
        elif line.startswith("origem:"):
            values['origem'] = line.split(":", 1)[1].strip()
        elif line.startswith("descrição:"):
            values['descricao'] = line.split(":", 1)[1].strip()
        elif line.startswith("revenda_cc:"):
            values['revenda_cc'] = line.split(":", 1)[1].strip()
        elif line.startswith("cc:"):
            values['cc'] = line.split(":", 1)[1].strip()
        elif line.startswith("rateio:"):
            values['rateio'] = line.split(":", 1)[1].strip()
        elif line.startswith("código de tributação:"):
            values['código de tributação'] = line.split(":", 1)[1].strip()

    return values

def save_attachment(part, directory):
    """Função para salvar o anexo em um diretório local com extensão .xml."""
    filename = decode_header_value(part.get_filename())
    
    if not filename:
        filename = 'untitled'
    
    # Garantir que a extensão seja .xml
    if not filename.lower().endswith('.xml'):
        filename += '.xml'
    
    # Criar diretório, se não existir
    if not os.path.exists(directory):
        os.makedirs(directory)

    filepath = os.path.join(directory, filename)

    with open(filepath, 'wb') as f:
        f.write(part.get_payload(decode=True))

    print(f'Anexo salvo em: {filepath}')

# Função para processar os centros de custo
def processar_centros_de_custo(cc_texto):
    """
    Processa o texto extraído para identificar se é valor ou porcentagem e retorna
    o formato adequado para o preenchimento.
    """
    centros_de_custo = []

    # Divida os centros de custo por vírgulas
    for item in cc_texto.split(','):
        item = item.strip()

        # Verifique se contém '%' indicando porcentagem
        if '%' in item:
            cc, porcentagem = item.split('-')
            cc = cc.strip()
            porcentagem = porcentagem.strip()
            centros_de_custo.append((cc, None, porcentagem))  # Adiciona como porcentagem
        else:
            cc, valor = item.split('-')
            cc = cc.strip()
            valor = valor.strip()
            centros_de_custo.append((cc, valor, None))  # Adiciona como valor monetário

    return centros_de_custo

# Código de automação

try:
    # Conectar ao servidor POP3 com SSL
    server = poplib.POP3_SSL(HOST, PORT)

    # Fazer login
    server.user(USERNAME)
    server.pass_(PASSWORD)

    # Obter o número de mensagens no servidor
    num_messages = len(server.list()[1])
    print(f'Você tem {num_messages} mensagem(s) no servidor.')

    # Percorrer todas as mensagens
    for i in range(num_messages):
        # Recuperar a mensagem (começando da mais recente)
        response, lines, octets = server.retr(i + 1)

        # Decodificar a mensagem
        raw_message = b'\n'.join(lines).decode('utf-8', errors='ignore')
        email_message = parser.Parser().parsestr(raw_message)

        # Decodificar o assunto
        subject = decode_header_value(email_message['subject'])

        # Verificar se o assunto corresponde ao ASSUNTO_ALVO
        if subject == ASSUNTO_ALVO:
            print(f"\nE-mail encontrado com o assunto: {subject}")

            # Obter o corpo do e-mail e anexos
            body = ""
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        charset = part.get_content_charset()
                        body = decode_body(part.get_payload(decode=True), charset)
                    else:  # Anexos
                        if part.get('Content-Disposition') is not None:
                            save_attachment(part, DIRECTORY)
            else:
                charset = email_message.get_content_charset()
                body = decode_body(email_message.get_payload(decode=True), charset)

            # Extrair valores
            valores_extraidos = extract_values(body)

            # Atribuir valores às variáveis
            departamento = valores_extraidos['departamento']
            origem = valores_extraidos['origem']
            descricao = valores_extraidos['descricao']
            revenda_cc = valores_extraidos['revenda_cc']
            cc_texto = valores_extraidos['cc']
            rateio = valores_extraidos['rateio']
            cod_item = valores_extraidos['código de tributação']

            print(f"Departamento: {departamento}")
            print(f"Origem: {origem}")
            print(f"Descrição: {descricao}")
            print(f"Revenda CC: {revenda_cc}")
            print(f"Centros de Custo: {cc_texto}")
            print(f"Rateio: {rateio}")
            print(f"Código de Tributação: {cod_item}")

            # Processar centros de custo
            centros_de_custo = processar_centros_de_custo(cc_texto)
            
            # Lançar os centros de custo
            for idx, (cc, valor, porcentagem) in enumerate(centros_de_custo):
                # Se não for o primeiro centro de custo, clique para adicionar um novo
                if idx > 0:
                    gui.click(1126, 758)  # Clica para adicionar novo centro de custo
                    time.sleep(1)  # Espera o sistema carregar o novo centro de custo
                    gui.write(revenda_cc) # Aqui várialvel relativa Revenda centro de custo puxar email
                    gui.press("tab", presses=3)
                    gui.write(cc)
                    gui.press("tab", presses=2)
                    gui.write(origem) # Aqui várialvel relativa de Origem puxar email
                    gui.press("tab", presses=2)
                
                # Preenche o valor ou porcentagem conforme o caso
                if valor is not None:
                    print(f"Centro de Custo: {cc}, Valor: {valor}")
                    gui.write(valor)  # Preenche o campo de valor
                if porcentagem is not None:
                    print(f"Centro de Custo: {cc}, Porcentagem: {porcentagem}")
                    gui.write(porcentagem)  # Preenche o campo de porcentagem
                gui.press("tab")  # Avança para o próximo campo, se necessário

            break  # Para após encontrar o primeiro e-mail com o assunto desejado

    # Desconectar do servidor
    server.quit()

except Exception as e:
    print('Erro:', e)

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

time.sleep(3)

# Localiza a janela do BRAVOS pelo título
window = gw.getWindowsWithTitle('BRAVOS v5.17 Evolutivo')[0]  # Assumindo que é a única com "BRAVOS" no título

# Centraliza a janela se necessário
window.activate()

# Calcula a posição relativa do ícone na barra de ferramentas
x, y = window.left + 275, window.top + 80  # Ajuste os offsets conforme necessário
time.sleep(3)
gui.moveTo(x, y, duration=0.5)
gui.click()
time.sleep(5)
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
gui.write(departamento)
gui.press("tab", presses=2)
gui.write(origem)
gui.press("tab", presses=19)
gui.write(chave_acesso)
gui.press("tab")
gui.write(modelo)
gui.press("tab", presses=18)
gui.press("right", presses=2)
gui.press("tab", presses=5)
gui.press("enter")
gui.press("tab", presses=4)
gui.write(cod_item) # variavel via email "Outros"
gui.press("tab", presses=10)
gui.write("1")
gui.press("tab")
gui.write(valor_total)
gui.press("tab", presses=26)
gui.write(descricao) # Variavel Relativa puxar pelo email
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
gui.write(revenda_cc) # Aqui várialvel relativa Revenda centro de custo puxar email
gui.press("tab", presses=3)
gui.write(cc)
gui.press("tab", presses=2)
gui.write(origem) # Aqui várialvel relativa de Origem puxar email
gui.press("tab", presses=2)
for cc, valor, porcentagem in centros_de_custo:
    gui.press("tab")
    if valor:
        gui.write(valor)  # Preenche o campo de valor
    if porcentagem:
        gui.write(porcentagem)  # Preenche o campo de porcentagem
    gui.press("tab")
    gui.moveTo(1202, 758)  
gui.press("tab")
gui.click()
gui.moveTo(1271, 758)
gui.click()
gui.press("tab", presses=3)
# gui.press("enter") Se tirar do comentario ele grava a nota no sistema