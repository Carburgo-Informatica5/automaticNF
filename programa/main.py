import os
import xml.etree.ElementTree as ET
import sys
import time
from datetime import datetime, timedelta
import pyautogui as gui
import pygetwindow as gw
import pytesseract
from email import parser
from email.header import decode_header
import poplib
import json
import logging
import unicodedata
import re
from logging.handlers import TimedRotatingFileHandler
from email.utils import parseaddr

from processar_xml import *
from db_connection import *
from DANImail import Queue, WriteTo

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/bravos")

from bravos import openBravos

logging.info("Iniciando o Programa")

HOST = "mail.carburgo.com.br"  # Servidor POP3 corporativo
PORT = 995  # Porta segura (SSL)
USERNAME = "dani@carburgo.com.br"
PASSWORD = "p@r!sA1856"
# Assunto alvo para busca
ASSUNTO_ALVO = "lançamento notas fiscais DANI"
# Diretório para salvar os anexos
DIRECTORY = os.path.join(current_dir, "anexos")


def decode_header_value(header_value):
    decoded_fragments = decode_header(header_value)
    decoded_string = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_string += fragment.decode(encoding or "utf-8")
        else:
            decoded_string += fragment
    return decoded_string


def normalize_text(text):
    return "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    ).lower()


def decode_body(payload, charset):
    if charset is None:
        charset = "utf-8"
    try:
        return payload.decode(charset)
    except (UnicodeDecodeError, LookupError):
        return payload.decode("ISO-8859-1", errors="ignore")


config_path = os.path.join(current_dir, "config.yaml")
with open(config_path, "r") as file:
    config = yaml.safe_load(file)
dani = Queue(config)


def extract_values(text):
    values = {
        "departamento": None,
        "origem": None,
        "descricao": None,
        "cc": None,
        "rateio": None,
        "cod_item": None,
    }
    lines = text.lower().splitlines()
    for line in lines:
        if line.startswith("departamento:"):
            values["departamento"] = line.split(":", 1)[1].strip()
        elif line.startswith("origem:"):
            values["origem"] = line.split(":", 1)[1].strip()
        elif line.startswith("descrição:"):
            values["descricao"] = line.split(":", 1)[1].strip()
        elif line.startswith("cc:"):
            values["cc"] = line.split(":", 1)[1].strip()
            logging.info(f"CC extraído: {values['cc']}")
        elif line.startswith("rateio:"):
            values["rateio"] = line.split(":", 1)[1].strip()
        elif line.startswith("código de tributação:"):
            values["cod_item"] = line.split(":", 1)[1].strip()
    return values


PROCESSED_EMAILS_FILE = os.path.join(current_dir, "processed_emails.json")

def load_processed_emails():
    if os.path.exists(PROCESSED_EMAILS_FILE):
        with open(PROCESSED_EMAILS_FILE, "r") as file:
            try:
                return json.load(file)
            except json.JSONDecodeError:
                logging.error("Erro ao carregar processed_emails.json. O arquivo está vazio ou corrompido.")
                return []
    return []

def save_processed_emails(processed_emails):
    with open(PROCESSED_EMAILS_FILE, "w") as file:
        json.dump(processed_emails, file)

def check_emails(nmr_nota):
    sender = None  # Inicializar a variável sender
    try:
        server = poplib.POP3_SSL(HOST, PORT)
        server.user(USERNAME)
        server.pass_(PASSWORD)
        num_messages = len(server.list()[1])
        logging.info(f"Conectado ao servidor POP3. Número de mensagens: {num_messages}")

        if not os.path.exists(DIRECTORY):
            os.makedirs(DIRECTORY)

        processed_emails = load_processed_emails()

        for i in range(num_messages):
            response, lines, octets = server.retr(i + 1)
            raw_message = b"\n".join(lines).decode("utf-8", errors="ignore")
            email_message = parser.Parser().parsestr(raw_message)
            subject = decode_header_value(email_message["subject"])
            from_header = decode_header_value(email_message["from"])
            sender = parseaddr(from_header)[1]
            email_id = email_message["Message-ID"]

            if email_id in processed_emails:
                logging.info(f"E-mail já processado: {subject}")
                continue

            logging.info(f"Verificando e-mail com assunto: {subject}")

            normalized_subject = normalize_text(subject)
            normalized_assunto_alvo = normalize_text(ASSUNTO_ALVO)

            if not subject:
                logging.error("Assunto do e-mail está vazio.")
                continue

            if normalized_subject == normalized_assunto_alvo:
                logging.info(f"E-mail encontrado com o assunto: {subject}")

                if email_message.is_multipart():
                    for part in email_message.walk():
                        content_type = part.get_content_type()
                        logging.info(
                            f"Parte do e-mail com tipo de conteúdo: {content_type}"
                        )
                        if content_type == "text/plain":
                            charset = part.get_content_charset()
                            body = decode_body(part.get_payload(decode=True), charset)
                        elif part.get("Content-Disposition") is not None:
                            logging.info("Encontrado anexo no e-mail")
                            dados_nota_fiscal = save_attachment(part, DIRECTORY)
                            if dados_nota_fiscal:
                                try:
                                    valores_extraidos = extract_values(body)
                                    departamento = valores_extraidos["departamento"]
                                    origem = valores_extraidos["origem"]
                                    descricao = valores_extraidos["descricao"]
                                    cc_texto = valores_extraidos["cc"]
                                    logging.info(
                                        f"Texto de centros de custo extraído: {cc_texto}"
                                    )
                                    rateio = valores_extraidos["rateio"]
                                    cod_item = valores_extraidos["cod_item"]
                                    if not dados_nota_fiscal["valor_total"]:
                                        raise ValueError(
                                            "Valor total não encontrado na nota fiscal"
                                        )
                                    valor_total = str(
                                        dados_nota_fiscal["valor_total"][0]["valor_total"]
                                    ).replace(".", ",")
                                    dados_centros_de_custo = process_cost_centers(
                                        cc_texto, float(valor_total.replace(",", "."))
                                    )
                                    logging.info(
                                        f"Dados dos centros de custo: {dados_centros_de_custo}"
                                    )
                                    dados_email = {
                                        "departamento": departamento,
                                        "origem": origem,
                                        "descricao": descricao,
                                        "cc": cc_texto,
                                        "cod_item": cod_item,
                                        "valor_total": valor_total,
                                        "dados_centros_de_custo": dados_centros_de_custo,
                                        "emitente": dados_nota_fiscal["emitente"],
                                        "num_nota": dados_nota_fiscal["num_nota"],
                                        "data_emi": dados_nota_fiscal["data_emi"],
                                        "data_venc": dados_nota_fiscal["data_venc"],
                                        "chave_acesso": dados_nota_fiscal["chave_acesso"],
                                        "modelo": dados_nota_fiscal["modelo"],
                                        "destinatario": dados_nota_fiscal["destinatario"],
                                        "rateio": rateio,
                                        "sender": sender,
                                        "email_id": email_id,  # Adiciona o ID do e-mail
                                        "parcelas": dados_nota_fiscal["pagamento_parcelado"],
                                        "serie": dados_nota_fiscal["serie"],
                                    }
                                    
                                    logging.info(f"Dados carregados: {dados_nota_fiscal}")
                                    logging.info(f"Parcelas carregadas: {dados_nota_fiscal.get('pagamento_parcelado')}")

                                    
                                    server.quit()
                                    return dados_email  # Retorna os dados do primeiro e-mail encontrado
                                except Exception as e:
                                    logging.error(f"Erro ao processar o e-mail: {e}")
                                    send_email_error(
                                        dani, sender, f"Erro ao processar o e-mail: {e}", nmr_nota,
                                    )
                                    server.quit()
                                    return None
                            else:
                                logging.error("Erro ao salvar ou processar o anexo")
                                send_email_error(
                                    dani, sender, "Erro ao salvar ou processar o anexo", nmr_nota,
                                )
                                server.quit()
                                return None
                else:
                    charset = email_message.get_content_charset()
                    body = decode_body(email_message.get_payload(decode=True), charset)

        server.quit()
        logging.info("Nenhum e-mail com o assunto alvo encontrado")
        return None
    except Exception as e:
        logging.error(f"Erro ao verificar emails: {e}")
        send_email_error(
            dani,
            sender if sender else "caetano.apollo@carburgo.com.br",
            f"Erro ao verificar emails: {e}",
            nmr_nota,
        )
        return None

def save_attachment(part, directory):
    filename = decode_header_value(part.get_filename())
    if not filename:
        filename = "untitled.xml"
    elif not filename.lower().endswith(".xml"):
        logging.info(f"Ignorando anexo não XML: {filename}")
        return None  # Ignorar anexos que não são XML
    content_type = part.get_content_type()
    logging.info(f"Tipo de conteúdo do anexo: {content_type}")
    if not content_type == "application/xml" and not filename.endswith(".xml"):
        logging.info(f"O anexo não é um arquivo XML: {filename}")
        return None
    if not os.path.exists(directory):
        os.makedirs(directory)
    filepath = os.path.join(directory, filename)
    logging.info(f"Salvando anexo em: {filepath}")
    with open(filepath, "wb") as f:
        f.write(part.get_payload(decode=True))
    logging.info(f"Anexo salvo em: {filepath}")
    dados_nota_fiscal = parse_nota_fiscal(filepath)
    if dados_nota_fiscal:
        salvar_dados_em_arquivo(dados_nota_fiscal, filename, "NOTAS EM JSON")
        return dados_nota_fiscal
    else:
        logging.error(f"Erro ao parsear o XML da nota fiscal: {filepath}")
        return None



def process_cost_centers(cc_texto, valor_total):
    valor_total = float(valor_total)
    centros_de_custo = []
    total_calculado = 0.0

    cc_texto = re.sub(r"[\u2013\u2014-]+", "-", cc_texto)

    if cc_texto.strip().isdigit():
        cc_texto = f"{cc_texto.strip()}-100%"

    itens = [item.strip() for item in cc_texto.split(",")]

    for i, item in enumerate(itens):
        try:
            cc, valor = [x.strip() for x in item.split("-")]
            logging.info(f"Processando centro de custo: {cc}, valor: {valor}")
            if "%" in valor:
                porcentagem = float(valor.replace("%", ""))
                valor_calculado = round((valor_total * porcentagem) / 100, 2)
            else:
                valor_calculado = float(valor.replace(",", "."))

            total_calculado += valor_calculado
            centros_de_custo.append(
                (cc, valor_calculado)
            )  # Mantém valor como float temporariamente
        except ValueError:
            logging.error(f"Erro ao processar centro de custo: {item}")

    diferenca = round(valor_total - total_calculado, 2)

    if abs(diferenca) > 0:
        if centros_de_custo:
            ultimo_cc, ultimo_valor = centros_de_custo[-1]
            # Converte o último valor para float, se necessário
            if isinstance(ultimo_valor, str):
                ultimo_valor = float(ultimo_valor.replace(",", "."))
            diferenca_float = ultimo_valor + diferenca  # Soma corretamente
            diferenca_str = f"{diferenca_float:.2f}".replace(
                ".", ","
            )  # Converte para string
            centros_de_custo[-1] = (ultimo_cc, diferenca_str)
        else:
            logging.error("Nenhum centro de custo encontrado para aplicar a diferença")
            raise ValueError(
                "Erro: Nenhum centro de custo encontrado para aplicar a diferença"
            )

    return centros_de_custo


def send_email_error(dani, destinatario, erro, nmr_nota):
    config["to"] = destinatario
    dani = Queue(config)

    mensagem = (
        dani.make_message()
        .set_color("red")
        .add_text(f"Erro durante lançamento de nota fiscal: {nmr_nota}", tag="h1")
        .add_text(str(erro), tag="pre")
    )

    mensagem_assinatura = (
        dani.make_message()
        .set_color("green")
        .add_text("DANI", tag="h1")
        .add_text("Este é um email enviado automaticamente pelo sistema DANI.", tag="p")
        .add_text(
            'Para reportar erros, envie um email para: ticket.dani@carburgo.com.br', 
            tag="p"
        )
        .add_text(
            'Em caso de dúvidas, entre em contato com: suporte.dani@carburgo.com.br',
            tag="p"
        )
    )
    dani.push(mensagem).push(mensagem_assinatura).flush()
    sys.exit(1)


def send_success_message(dani, destinatario, nmr_nota):
    config["to"] = destinatario
    dani = Queue(config)

    mensagem = (
        dani.make_message()
        .set_color("green")
        .add_text(f"Nota lançada com sucesso número da nota: {nmr_nota}", tag="h1")
        .add_text("Acesse o sistema para verificar o lançamento.", tag="pre")
    )

    mensagem_assinatura = (
        dani.make_message()
        .set_color("green")
        .add_text("DANI", tag="h1")
        .add_text("Este é um email enviado automaticamente pelo sistema DANI.", tag="p")
        .add_text(
            'Para reportar erros, envie um email para: ticket.dani@carburgo.com.br', 
            tag="p"
        )
        .add_text(
            'Em caso de dúvidas, entre em contato com: caetano.apollo@carburgo.com.br',
            tag="p"
        )
    )

    dani.push(mensagem).push(mensagem_assinatura).flush()

def processar_parcelas(parcelas):
    logging.info("Iniciando o processamento das parcelas")
    for parcela in parcelas:
        logging.info(f"Processando parcela: {parcela}")

        data_vencimento = parcela["data_vencimento"]
        valor_parcela = parcela["valor_parcela"]

        # Formatar a data de vencimento
        data_vencimento_formatada = datetime.strptime(data_vencimento, "%Y-%m-%d")

        # Adicionar um dia à data de vencimento
        data_vencimento_ajustada = data_vencimento_formatada + timedelta(days=1)

        # Calcular os dias restantes para o vencimento ajustado
        dias_para_vencimento = (data_vencimento_ajustada - datetime.now()).days

        # Log para verificar a data ajustada e os dias calculados
        logging.info(
            f"Parcela ajustada para vencimento em {data_vencimento_ajustada.strftime('%d/%m/%Y')} "
            f"faltam {dias_para_vencimento} dias."
        )

        # Passar os dias ajustados para o sistema
        gui.write(str(dias_para_vencimento))
        gui.press("tab", presses=2)
        gui.write(valor_parcela.replace(".", ","))
        gui.press("enter")
        gui.press("enter")

dados_nota_fiscal = None

logging.info("Carregando dados da nota fiscal")
if dados_nota_fiscal is not None:
    nome_emitente = dados_nota_fiscal["emitente"]["nome"]
    cnpj_emitente = dados_nota_fiscal["emitente"]["cnpj"]
    nome_dest = dados_nota_fiscal["destinatario"]["nome"]
    cnpj_dest = dados_nota_fiscal["destinatario"]["cnpj"]
    chave_acesso = dados_nota_fiscal["chave_acesso"]["chave"]
    nmr_nota = dados_nota_fiscal["num_nota"]["numero_nota"]
    data_emi = dados_nota_fiscal["data_emi"]["data_emissao"]
    data_venc = dados_nota_fiscal["data_venc"]["data_venc"]
    modelo = dados_nota_fiscal["modelo"]["modelo"]
    parcelas = dados_nota_fiscal["pagamento_parcelado"]
else:
    logging.error("Dados da nota fiscal não foram carregados")


class SystemNF:
    logging.info("Entrou na classe do Sistema")

    def __init__(self, master=None):
        self.br = None

    def automation_gui(
        self,
        departamento,
        origem,
        descricao,
        cc,
        cod_item,
        valor_total,
        dados_centros_de_custo,
        cnpj_emitente,
        nmr_nota,
        data_emi,
        data_venc,
        chave_acesso,
        modelo,
        rateio,
        parcelas,
        serie,
    ):
        if not parcelas:
            logging.warning("As parcelas estão vazias ao entrar em automation_gui")
        logging.info("Entrou na parte da automação")
        logging.info(f"Departamento: {departamento}")
        logging.info(f"Origem: {origem}")
        logging.info(f"Descrição: {descricao}")
        logging.info(f"CC: {cc}")
        logging.info(f"Código do Item: {cod_item}")
        logging.info(f"Valor Total: {valor_total}")
        logging.info(f"Dados dos Centros de Custo: {dados_centros_de_custo}")
        logging.info(f"CNPJ Emitente: {cnpj_emitente}")
        logging.info(f"Número da Nota: {nmr_nota}")
        logging.info(f"Data de Emissão: {data_emi}")
        logging.info(f"Data de Vencimento: {data_venc}")
        logging.info(f"Chave de Acesso: {chave_acesso}")
        logging.info(f"Modelo: {modelo}")
        logging.info(f"Rateio: {rateio}")
        logging.info(f"Série: {serie}")
        logging.info(f"Parcelas: {parcelas}")

        try:
            data_atual = datetime.now()
            data_formatada = data_atual.strftime("%d%m%Y")
            time.sleep(3)
            window = gw.getWindowsWithTitle("BRAVOS v5.18 Evolutivo")[0]
            if not window:
                raise Exception("Janela do BRAVOS não encontrada")
            window.activate()
            x, y = window.left + 275, window.top + 80
            time.sleep(3)
            gui.moveTo(x, y, duration=0.5)
            gui.click()
            time.sleep(15)
            if empresa == 1:
                gui.press("tab", presses=22)
            else:
                gui.press("tab", presses=19)
            gui.write(cnpj_emitente)
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
            threshold = 190
            screenshot = screenshot.point(lambda p: p > threshold and 255)
            config = r"--psm 7 outputbase digits"
            cliente = pytesseract.image_to_string(screenshot, config=config)
            
            # Modelo 55, 43 e 22 não podem ser código de tributação nunca é dez
            
            time.sleep(5)
            gui.hotkey("ctrl", "f4")
            gui.press("alt")
            gui.press("right", presses=6)
            gui.press("down", presses=4)
            gui.press("enter")
            gui.press("down", presses=2)
            gui.press("enter")
            time.sleep(5)
            gui.press("tab")
            gui.write(nmr_nota)
            gui.press("tab")
            gui.write(serie) 
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
            if modelo == "55":
                gui.write(chave_acesso)
            gui.press("tab")
            gui.write(modelo)
            gui.press("tab", presses=18)
            
            # A partir daqui que deve ser implementado a nota de serviço
            
            # Se arquivo endswitch (.pdf)
            
            # Se arquivo endswitch (.xml)
            gui.press("right", presses=2)
            gui.press("tab", presses=5)
            gui.press("enter")
            gui.press("tab", presses=4)
            if (modelo == "55" or modelo == "43" or modelo == "22") and cod_item != "7":
                send_email_error(
                    dani, dados_email.get("sender", "caetano.apollo@carburgo.com.br"), "Erro: Modelo de nota fiscal inválido para o código de tributação informado.", nmr_nota, 
                )
            gui.write(cod_item)
            gui.press("tab", presses=10)
            gui.write("1")
            gui.press("tab")
            logging.info(f"Valor total: {valor_total}")
            gui.write(valor_total)
            gui.press("tab", presses=26)
            gui.write(descricao)
            gui.press(["tab", "enter"])
            gui.press("tab", presses=11)
            gui.press("left", presses=2)
            gui.press("tab", presses=5)
            gui.press("enter")
            
            logging.info(f"Processando parcelas: {parcelas}")
            gui.press("tab")
            gui.press("enter")
            processar_parcelas(parcelas)
            
            gui.press("tab", presses=3)
            gui.press(["enter", "tab", "tab", "tab", "enter"])
            gui.press("tab", presses=35)
            if rateio.lower() == "sim":
                gui.press("enter")
                gui.press("tab", presses=8)
                logging.info("Pressionou o tab corretamente")
                total_rateio = 0
                for i, (cc, valor) in enumerate(dados_centros_de_custo):
                    logging.info("Lançando centro de custo")
                    gui.write(str(revenda))
                    gui.press("tab", presses=3)
                    gui.write(cc)
                    gui.press("tab", presses=2)
                    gui.write(origem)
                    gui.press("tab", presses=2)
                    if i == len(dados_centros_de_custo) - 1:
                        valor = float(valor_total.replace(",", ".")) - total_rateio
                    else:
                        total_rateio += float(valor)
                    gui.write(f"{valor:.2f}".replace(".", ","))
                    gui.press("f2", interval=2)
                    gui.press("f3")
                    if i == len(dados_centros_de_custo) - 1:
                        gui.press("f2", interval=2)
                        gui.press("esc", presses=3)
                        logging.info("Último centro de custo salvo e encerrado.")
                    else:
                        gui.press("tab", presses=3)
            gui.press("tab", presses=4)
            gui.press("enter")

        except Exception as e:
            send_email_error(
                dani, dados_email.get("sender", "caetano.apollo@carburgo.com.br"), e, nmr_nota,
            )
            print(f"Erro durante a automação: {e}")
            print("Automação iniciada com os dados extraídos.")

if __name__ == "__main__":
    while True:
        logging.info("Iniciando a automação")
        try:
            nmr_nota = ''
            dados_email = check_emails(nmr_nota)
            if dados_email is not None:
                logging.info("Executando automação GUI")
                logging.info(f"Conteúdo completo de dados_email: {dados_email}")

                departamento = dados_email.get("departamento")
                origem = dados_email.get("origem")
                descricao = dados_email.get("descricao")
                cc = dados_email.get("cc")
                cod_item = dados_email.get("cod_item")
                valor_total = dados_email.get("valor_total")
                dados_centros_de_custo = dados_email.get("dados_centros_de_custo")
                rateio = dados_email.get("rateio")
                parcelas = dados_email.get("parcelas", [])
                serie = dados_email.get("serie")
                if "parcelas" in dados_email:
                    parcelas = dados_email["parcelas"]
                else:
                    logging.warning("Chave 'parcelas' não encontrada em dados_email")
                logging.info(f"Parcelas recebidas: {parcelas}")

                if (
                    "emitente" in dados_email
                    and "num_nota" in dados_email
                    and "data_emi" in dados_email
                    and "data_venc" in dados_email
                    and "chave_acesso" in dados_email
                    and "modelo" in dados_email
                    and "destinatario" in dados_email
                ):
                    cnpj_emitente = dados_email["emitente"]["cnpj"]
                    nmr_nota = dados_email["num_nota"]["numero_nota"]
                    data_emi = dados_email["data_emi"]["data_emissao"]
                    data_venc = dados_email["data_venc"]["data_venc"]
                    chave_acesso = dados_email["chave_acesso"]["chave"]
                    modelo = dados_email["modelo"]["modelo"]
                    cnpj_dest = dados_email["destinatario"]["cnpj"]
                else:
                    logging.error(
                        "Dados da nota fiscal não foram carregados corretamente"
                    )
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        "Erro: Dados da nota fiscal não foram carregados corretamente",
                        nmr_nota,
                    )
                    continue

                # Verificação de campos obrigatórios
                campos_obrigatorios = [
                    departamento,
                    origem,
                    descricao,
                    cc,
                    cod_item,
                    valor_total,
                    dados_centros_de_custo,
                ]
                if not all(campos_obrigatorios):
                    mensagem_erro = (
                        "Faltando campos obrigatórios para o lançamento:\n"
                    )
                    if not departamento:
                        mensagem_erro += "- Departamento\n"
                    if not origem:
                        mensagem_erro += "- Origem\n"
                    if not descricao:
                        mensagem_erro += "- Descrição\n"
                    if not cc:
                        mensagem_erro += "- CC\n"
                    if not cod_item:
                        mensagem_erro += "- Código de tributação do item\n"
                    if not valor_total:
                        mensagem_erro += "- Valor Total\n"
                    if not dados_centros_de_custo:
                        mensagem_erro += "- Dados dos Centros de Custo\n"
                    logging.error(mensagem_erro)
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        mensagem_erro,
                        nmr_nota,
                    )
                    continue

                # Executando a parte de revenda primeiro
                if cnpj_dest:
                    result = revenda(cnpj_dest)
                    if result:
                        empresa, revenda = result
                        logging.info(f"Empresa: {empresa}, Revenda: {revenda}")

                        # Acessando o menu
                        gui.press("alt")
                        gui.press("right")
                        gui.press("down")
                        gui.press("enter")
                        gui.press("down", presses=2)

                        gui.write(f"{empresa}.{revenda}")
                        gui.press("enter")
                        time.sleep(5)
                    else:
                        send_email_error(
                            dani,
                            dados_email.get(
                                "sender", "caetano.apollo@carburgo.com.br"
                            ),
                            "Erro, CNPJ do destinatário não encontrado",
                            nmr_nota,
                        )
                        logging.error("Erro, CNPJ não encontrado")

                # Adicionando logs para verificar os dados extraídos
                logging.info(f"Departamento: {departamento}")
                logging.info(f"Origem: {origem}")
                logging.info(f"Descrição: {descricao}")
                logging.info(f"CC: {cc}")
                logging.info(f"Código do Item: {cod_item}")
                logging.info(f"Valor Total: {valor_total}")
                logging.info(
                    f"Dados dos Centros de Custo: {dados_centros_de_custo}"
                )
                logging.info(f"CNPJ Eminente: {cnpj_emitente}")
                logging.info(f"Número da Nota: {nmr_nota}")
                logging.info(f"Data de Emissão: {data_emi}")
                logging.info(f"Data de Vencimento: {data_venc}")
                logging.info(f"Chave de Acesso: {chave_acesso}")
                logging.info(f"Modelo: {modelo}")

                sistema_nf = SystemNF()

                logging.info(f"Parcelas a serem passadas para automation_gui: {parcelas}")

                try:
                    sistema_nf.automation_gui(
                        departamento,
                        origem,
                        descricao,
                        cc,
                        cod_item,
                        valor_total,
                        dados_centros_de_custo,
                        cnpj_emitente,
                        nmr_nota,
                        data_emi,
                        data_venc,
                        chave_acesso,
                        modelo,
                        rateio,
                        parcelas,
                        serie, 
                    )
                    # Adicionar o ID do e-mail processado à lista após lançamento bem-sucedido
                    processed_emails = load_processed_emails()
                    processed_emails.append(dados_email["email_id"])
                    save_processed_emails(processed_emails)
                    send_success_message(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        nmr_nota,
                    )
                except Exception as e:
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        e,
                        nmr_nota,
                    )
            else:
                logging.info("Nenhum dado extraído, automação não será executada")
        except Exception as e:
            logging.error(f"Erro durante a automação: {e}")
            send_email_error(
                dani, "caetano.apollo@carburgo.com.br", e, nmr_nota,
            )
        logging.info("Esperando antes da nova verificação...")
        time.sleep(30)