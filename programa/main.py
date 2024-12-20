import os
import xml.etree.ElementTree as ET
import sys
import time
import datetime
import pyautogui as gui
import pygetwindow as gw
import pytesseract
from email import parser
from email.header import decode_header
import poplib
import json
import logging

from processar_xml import *

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(current_dir, ".."))

from bravos import openBravos

logging.info("Iniciando o Programa")

HOST = "mail.carburgo.com.br"  # Servidor POP3 corporativo
PORT = 995  # Porta segura (SSL)
USERNAME = "caetano.apollo@carburgo.com.br"
PASSWORD = "p@r!sA1856"
# Assunto alvo para busca
ASSUNTO_ALVO = "Lançamentos notas fiscais DANI"
# Diretório para salvar os anexos
DIRECTORY = "anexos"  # Pasta onde os anexos serão salvos


def decode_header_value(header_value):
    decoded, encoding = decode_header(header_value)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else "utf-8", errors="ignore")
    return decoded


def decode_body(payload, charset):
    if charset is None:
        charset = "utf-8"
    try:
        return payload.decode(charset)
    except (UnicodeDecodeError, LookupError):
        return payload.decode("ISO-8859-1", errors="ignore")


def extract_values(text):
    values = {
        "departamento": None,
        "origem": None,
        "descricao": None,
        "revenda_cc": None,
        "cc": None,
        "rateio": None,
        "cod_item": None,
    }
    lines = text.splitlines()
    for line in lines:
        if line.startswith("departamento:"):
            values["departamento"] = line.split(":", 1)[1].strip()
        elif line.startswith("origem:"):
            values["origem"] = line.split(":", 1)[1].strip()
        elif line.startswith("descrição:"):
            values["descricao"] = line.split(":", 1)[1].strip()
        elif line.startswith("revenda_cc:"):
            values["revenda_cc"] = line.split(":", 1)[1].strip()
        elif line.startswith("cc:"):
            values["cc"] = line.split(":", 1)[1].strip()
        elif line.startswith("rateio:"):
            values["rateio"] = line.split(":", 1)[1].strip()
        elif line.startswith("código de tributação:"):
            values["cod_item"] = line.split(":", 1)[1].strip()
    return values


logging.info("Carregou e processou o email")


def processar_centros_de_custo(cc_texto, valor_total):
    """
    Processa o texto extraído para identificar se é valor ou porcentagem e retorna
    o formato adequado para o preenchimento.
    """
    centros_de_custo = []
    total_calculado = 0.0

    # Divida os centros de custo por vírgulas
    for item in cc_texto.split(","):
        item = item.strip()
        try:
            cc, valor = [x.strip() for x in item.split("-")]
            if "%" in valor:
                porcentagem = float(valor.replace("%", ""))
                valor_calculado = round((valor_total * porcentagem) / 100, 2)
            else:
                valor_calculado = float(valor.replace(",", "."))

            total_calculado += valor_calculado
            centros_de_custo.append((cc, valor_calculado))
        except ValueError:
            print(f"Erro ao processar centro de custo: {item}")

    # Ajuste final para compensar arredondamentos
    diferenca = round(valor_total - total_calculado, 2)

    if abs(diferenca) > 0.01:
        raise ValueError("Erro: Diferença de cálculo muito grande!")

    if abs(diferenca) > 0:
        # Aplica a diferença ao último centro de custo
        ultimo_cc, ultimo_valor = centros_de_custo[-1]
        centros_de_custo[-1] = (ultimo_cc, round(ultimo_valor + diferenca, 2))

    return centros_de_custo

dados_nota_fiscal = None

for _ in iter(int, 1):
    try:
        server = poplib.POP3_SSL(HOST, PORT)
        server.user(USERNAME)
        server.pass_(PASSWORD)
        num_messages = len(server.list()[1])
        print(f"Conectado ao servidor POP3. Número de mensagens: {num_messages}")
        
        pasta_notas = "C:/Users/VAS MTZ/Desktop/Caetano Apollo/NOTA EM JSON"
        if not os.path.exists(pasta_notas):
            os.makedirs(pasta_notas)
        
        for arquivo in os.listdir(pasta_notas):
            if arquivo.endswith(".json"):
                caminho_completo = os.path.join(pasta_notas, arquivo)
                with open(caminho_completo, "r") as json_file:
                    dados_nota_fiscal = json.load(json_file)
    
        def save_attachment(part, directory):
            filename = decode_header_value(part.get_filename())
            if not filename:
                filename = "untitled.xml"
            elif not filename.lower().endswith(".xml"):
                filename += ".xml"
            content_type = part.get_content_type()
            print(f"Tipo de conteúdo do anexo: {content_type}")
            if not content_type == "application/xml" and not filename.endswith(".xml"):
                print(f"O anexo não é um arquivo XML: {filename}")
                return
            if not os.path.exists(directory):
                os.makedirs(directory)
            filepath = os.path.join(directory, filename)
            print(f"Salvando anexo em: {filepath}")
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            logging.info(f"Anexo salvo em: {filepath}")
            dados_nota_fiscal = parse_nota_fiscal(filepath)
            if dados_nota_fiscal:
                salvar_dados_em_arquivo(dados_nota_fiscal, filename, "NOTA EM JSON")
                return dados_nota_fiscal
    
        # Dentro do loop de leitura dos e-mails
        for i in range(num_messages):
            response, lines, octets = server.retr(i + 1)
            raw_message = b"\n".join(lines).decode("utf-8", errors="ignore")
            email_message = parser.Parser().parsestr(raw_message)
            subject = decode_header_value(email_message["subject"])
            print(f"Verificando e-mail com assunto: {subject}")
            
            if subject == ASSUNTO_ALVO:
                print(f"E-mail encontrado com o assunto: {subject}")
                
                if email_message.is_multipart():
                    for part in email_message.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/plain":
                            charset = part.get_content_charset()
                            body = decode_body(part.get_payload(decode=True), charset)
                        elif part.get("Content-Disposition") is not None:
                            dados_nota_fiscal = save_attachment(part, DIRECTORY)
                            if dados_nota_fiscal:  # Se o processamento da nota fiscal for bem-sucedido
                                break
                else:
                    charset = email_message.get_content_charset()
                    body = decode_body(email_message.get_payload(decode=True), charset)
                
                if dados_nota_fiscal is not None:  # Verifique se dados_nota_fiscal foi preenchido
                    valores_extraidos = extract_values(body)
                    departamento = valores_extraidos["departamento"]
                    origem = valores_extraidos["origem"]
                    descricao = valores_extraidos["descricao"]
                    revenda_cc = valores_extraidos["revenda_cc"]
                    cc_texto = valores_extraidos["cc"]
                    logging.info(f"Texto de centros de custo extraído: {cc_texto}")
                    rateio = valores_extraidos["rateio"]
                    cod_item = valores_extraidos["cod_item"]
                    valor_total = str(dados_nota_fiscal["valor_total"][0]["valor_total"]).replace(".", ",")
                    dados_centros_de_custo = processar_centros_de_custo(cc_texto, float(valor_total.replace(",", ".")))
                    logging.info(f"Dados dos centros de custo: {dados_centros_de_custo}")
                    
                    
                    sistema = SistemaNF()
                    sistema.executar_automacao_gui()
                    server.dele(i + 1)
                    break
                else:
                    logging.error("Não foi possível processar os dados da nota fiscal")
            else:
                logging.info(f"E-mail com assunto diferente: {subject}")
        
        # Continuar o código abaixo, verificando se dados_nota_fiscal foi carregado corretamente
        if dados_nota_fiscal is not None:
            try:
                nome_eminente = dados_nota_fiscal["eminente"]["nome"]
                cnpj_eminente = dados_nota_fiscal["eminente"]["cnpj"]
                nome_dest = dados_nota_fiscal["destinatario"]["nome"]
                cnpj_dest = dados_nota_fiscal["destinatario"]["cnpj"]
                chave_acesso = dados_nota_fiscal["chave_acesso"]["chave"]
                nmr_nota = dados_nota_fiscal["num_nota"]["numero_nota"]
                data_emi = dados_nota_fiscal["data_emi"]["data_emissao"]
                data_venc = dados_nota_fiscal["valor_total"][0]["data_venc"]
                modelo = dados_nota_fiscal["modelo"]["modelo"]
                logging.info("Dados da nota fiscal carregados")
            except KeyError as e:
                logging.error(f"Erro ao acessar dados da nota fiscal: {e}")
        else:
            logging.error("dados_nota_fiscal não foi carregado.")
        
        server.quit()
    except Exception as e:
        print(f"Erro: {e}")
    
    logging.info("Carregando dados da nota fiscal")
    
    # Continuar o processamento dos dados da nota fiscal após o loop
    if dados_nota_fiscal is not None:
        nome_eminente = dados_nota_fiscal["eminente"]["nome"]
        cnpj_eminente = dados_nota_fiscal["eminente"]["cnpj"]
        nome_dest = dados_nota_fiscal["destinatario"]["nome"]
        cnpj_dest = dados_nota_fiscal["destinatario"]["cnpj"]
        chave_acesso = dados_nota_fiscal["chave_acesso"]["chave"]
        nmr_nota = dados_nota_fiscal["num_nota"]["numero_nota"]
        data_emi = dados_nota_fiscal["data_emi"]["data_emissao"]
        data_venc = dados_nota_fiscal["valor_total"][0]["data_venc"]
        modelo = dados_nota_fiscal["modelo"]["modelo"]
    
    logging.info("Dados da nota fiscal carregados")


    class SistemaNF:
        logging.info("Entrou na classe do Sistema")

        def __init__(self, master=None):
            self.br = None

        def executar_automacao_gui(self):
            logging.info("Entrou na parte da automação")
            try:
                gui.PAUSE = 0.5
                data_atual = datetime.datetime.now()

                data_formatada = data_atual.strftime("%d%m%Y")

                gui.alert(
                    "O código vai começar. Não utilize nada do computador até o código finalizar!"
                )

                time.sleep(3)

                # Localiza a janela do BRAVOS pelo título
                window = gw.getWindowsWithTitle("BRAVOS v5.17 Evolutivo")[
                    0
                ]  # Assumindo que é a única com "BRAVOS" no título

                # Centraliza a janela se necessário
                window.activate()

                # Calcula a posição relativa do ícone na barra de ferramentas
                x, y = (
                    window.left + 275,
                    window.top + 80,
                )  # Ajuste os offsets conforme necessário
                time.sleep(3)
                gui.moveTo(x, y, duration=0.5)
                gui.click()
                time.sleep(15)
                gui.press("tab", presses=19)
                gui.write(cnpj_eminente)
                time.sleep(2)
                gui.press("enter")

                pytesseract.pytesseract_cmd = (
                    r"C:\Program Files\Tesseract-OCR/tesseract.exe"
                )

                nova_janela = gw.getActiveWindow()

                janela_left = nova_janela.left
                janela_top = nova_janela.top

                time.sleep(5)

                x, y, width, height = janela_left + 500, janela_top + 323, 120, 21

                screenshot = gui.screenshot(region=(x, y, width, height))

                screenshot = screenshot.convert("L")
                threshold = 150
                screenshot = screenshot.point(lambda p: p > threshold and 255)

                config = r"--psm 7 outputbase digits"

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
                gui.write("1")
                gui.press("tab")
                gui.press("down")
                gui.press("tab")
                gui.write("0")
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
                gui.write(cod_item)  # variavel via email "Outros"
                gui.press("tab", presses=10)
                gui.write("1")
                gui.press("tab")
                logging.info(f"Valor total: {valor_total}")
                gui.write(valor_total)
                gui.press("tab", presses=26)
                gui.write(descricao)  # Variavel Relativa puxar pelo email
                gui.press(["tab", "enter"])
                gui.press("tab", presses=11)
                gui.press("left", presses=2)
                gui.press("tab", presses=5)
                gui.press(["enter", "tab", "enter"])
                gui.press("tab", presses=5)
                gui.write(data_venc)
                gui.press("tab", presses=4)
                gui.press(["enter", "tab", "tab", "tab", "enter"])
                gui.press("tab", presses=36)
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.hotkey("ctrl", "del")
                gui.press("enter")
                gui.press("tab", presses=8)
                logging.info("Pressionou o tab corretamente")
                for i, (cc, valor) in enumerate(dados_centros_de_custo):
                    logging.info("Lançando centro de custo")
                    gui.write(revenda_cc)
                    gui.press("tab", presses=3)
                    gui.write(cc)
                    gui.press("tab", presses=2)
                    gui.write(origem)
                    gui.press("tab", presses=2)
                    gui.write(f"{valor:.2f}".replace(".", ","))
                    gui.press("f2", interval=2)
                    gui.press("f3")

                    if i == len(dados_centros_de_custo) - 1:
                        gui.press("esc", presses=2)
                        logging.info("Último centro de custo salvo e encerrado.")
                    else:
                        # Avança para o próximo centro de custo
                        gui.press("tab", presses=3)

                gui.press("tab", presses=3)
                gui.press("enter")
            except Exception as e:
                print(f"Erro durante a automação: {e}")

                print("Automação iniciada com os dados extraídos.")
            except Exception as e:
                print(f"Erro durante a automação: {e}")


# if __name__ == "__main__":
#     sistema = SistemaNF()
#     sistema.executar_automacao_gui()