import os
import xml.etree.ElementTree as ET
import sys
import time
from datetime import datetime, timedelta
import calendar
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
import yaml
from typing import Callable, Any

from email.utils import parseaddr

from processar_xml import *
from db_connection import revenda, login as db_login
from DANImail import Queue, WriteTo
from gemini_api import GeminiAPI
from gemini_main import *

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/bravos")

from bravos.infoBravos import bravos as BravosClass

logging.info("Iniciando o Programa")

HOST = "mail.carburgo.com.br"  # Servidor POP3 corporativo
PORT = 995  # Porta segura (SSL)
USERNAME = "dani@carburgo.com.br"
PASSWORD = "p@r!sA1856"
# Assunto alvo para busca
if "ASSUNTO_ALVO" not in globals():
    ASSUNTO_ALVO = "lançamento frete"
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
    if not isinstance(text, str):  # Garante que text seja string
        logging.error(f"Erro: normalize_text recebeu {type(text)} em vez de string.")
        return ""
    
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
    if not isinstance(text, str): 
        logging.error(f"Erro: extract_values recebeu {type(text)} em vez de string.")
        return {}

    values = {
        "departamento": None,
        "origem": None,
        "descricao": None,
        "cc": None,
        "rateio": None,
        "cod_item": None,
        "data_vencimento": None,
        "tipo_imposto": None,
        "senha_arquivo": None,
        "senha_user": None
    }

    lines = text.lower().splitlines()
    for line in lines:
        if isinstance(line, str): 
            if line.startswith("departamento:"):
                values["departamento"] = line.split(":", 1)[1].strip()
            elif line.startswith("origem:"):
                values["origem"] = line.split(":", 1)[1].strip()
            elif line.startswith("descrição:"):
                values["descricao"] = line.split(":", 1)[1].strip()
            elif line.startswith("cc:"):
                values["cc"] = line.split(":", 1)[1].strip()
            elif line.startswith("rateio:"):
                values["rateio"] = line.split(":", 1)[1].strip()
            elif line.startswith("código de tributação:"):
                values["cod_item"] = line.split(":", 1)[1].strip()
            
            #somente para PDF
            elif "data vencimento" in line.lower():
                data_vencimento = line.split(":", 1)[1].strip()
                values["data_vencimento"] = data_vencimento.replace("/", "")
            elif line.startswith("tipo imposto:"):
                values["tipo_imposto"] = line.split(":", 1)[1].strip()
            elif line.startswith("senha arquivo:"):
                values["senha_arquivo"] = line.split(":", 1)[1].strip()
            elif line.startswith("modelo:"):
                values["modelo_email"] = line.split(":", 1)[1].strip()
            elif line.startswith("tipo documento:"):
                values["tipo_documento"] = line.split(":", 1)[1].strip()
            elif line.startswith("senha usuário:"):
                values["senha_user"] = line.split(":", 1)[1].strip()

    logging.info(f"Valores extraídos: {values}")

    return values

PROCESSED_EMAILS_FILE = os.path.join(current_dir, "processed_emails.json")

def localizar_e_clicar_no_campo(imagem_label, offset_x=150, confidence=0.9):
    """
    Localiza um rótulo na tela usando uma imagem e clica no campo ao lado.

    Args:
        imagem_label (str): Caminho para o arquivo de imagem do rótulo.
        offset_x (int): Distância em pixels para a direita do centro do rótulo
                        onde o clique deve ocorrer. Ajuste este valor conforme necessário.
        confidence (float): Nível de confiança para a busca da imagem (entre 0.0 e 1.0).

    Returns:
        bool: True se o campo foi encontrado e clicado, False caso contrário.
    """
    try:
        logging.info(f"Procurando pelo rótulo: {imagem_label}")
        # Tenta localizar a imagem do rótulo na tela
        posicao_label = gui.locateOnScreen(imagem_label, confidence=confidence)

        if posicao_label:
            # Se encontrou, calcula as coordenadas do centro do rótulo
            label_centro_x, label_centro_y = gui.center(posicao_label)
            logging.info(f"Rótulo '{imagem_label}' encontrado em: ({label_centro_x}, {label_centro_y})")

            # Calcula a posição do campo (à direita do rótulo)
            campo_x = label_centro_x + offset_x
            campo_y = label_centro_y

            # Clica no campo e o ativa
            logging.info(f"Clicando no campo de faturamento em: ({campo_x}, {campo_y})")
            gui.click(campo_x, campo_y)
            time.sleep(0.5) # Pequena pausa para o sistema responder
            return True
        else:
            logging.error(f"Rótulo '{imagem_label}' não encontrado na tela.")
            return False
    except FileNotFoundError:
        logging.error(f"Arquivo de imagem '{imagem_label}' não encontrado. Verifique o caminho.")
        return False
    except Exception as e:
        # Captura outras exceções do pyautogui (ex: imagem não encontrada na tela)
        logging.error(f"Ocorreu um erro ao tentar localizar a imagem: {e}")
        return False

def load_processed_emails():
    if os.path.exists(PROCESSED_EMAILS_FILE):
        with open(PROCESSED_EMAILS_FILE, "r") as file:
            try:
                return json.load(file)
            except json.JSONDecodeError:
                logging.error(
                    "Erro ao carregar processed_emails.json. O arquivo está vazio ou corrompido."
                )
                return []
    return []


def save_processed_emails(processed_emails):
    with open(PROCESSED_EMAILS_FILE, "w") as file:
        json.dump(processed_emails, file)

def processar_email_frete(email_message):
    logging.info("="*50)
    logging.info("INICIANDO PROCESSAMENTO DE E-MAIL DE FRETE")
    logging.info("="*50)
    
    sender = None
    dados_email = {}
    nmr_nota = "N/A"

    try:
        from_header = decode_header_value(email_message["from"])
        sender = parseaddr(from_header)[1]
        email_id = email_message["Message-ID"]

        processed_emails = load_processed_emails()
        if email_id in processed_emails:
            logging.warning(f"E-mail (ID: {email_id}) já processado. Ignorando.")
            return

        # 1. Extrair corpo e dados do e-mail
        body = ""
        if email_message.is_multipart():
            for part in email_message.walk():
                if part.get_content_type() == "text/plain" and "attachment" not in (part.get("Content-Disposition") or ""):
                    body = decode_body(part.get_payload(decode=True), part.get_content_charset())
                    break
        else:
            body = decode_body(email_message.get_payload(decode=True), email_message.get_content_charset())
        
        dados_email = extract_values(body)
        dados_email["body"] = body

        # 2. Validar usuário e senha
        resultado_login_db = db_login(sender)
        if not resultado_login_db:
            raise ValueError(f"Usuário {sender} não encontrado no banco de dados.")

        usuario_login = resultado_login_db[0] if isinstance(resultado_login_db, tuple) else resultado_login_db
        senha_login_from_email = dados_email.get("senha_user")
        if not senha_login_from_email:
            raise ValueError("Senha do usuário ('senha usuário:') não encontrada no corpo do e-mail.")
        
        dados_email["usuario_login"] = usuario_login
        dados_email["senha_login"] = senha_login_from_email
        dados_email["sender"] = sender
        dados_email["email_id"] = email_id

        # 3. Processar Anexo
        attachment_processed = False
        for part in email_message.walk():
            if part.get("Content-Disposition") and "attachment" in part.get("Content-Disposition"):
                logging.info("Anexo encontrado. Processando...")
                dados_anexo = save_attachment(part, DIRECTORY, dados_email)
                if not dados_anexo or "json_path" not in dados_anexo:
                    raise ValueError("Falha ao salvar, processar ou gerar JSON do anexo.")
                
                with open(dados_anexo["json_path"], "r", encoding="utf-8") as f:
                    json_data = json.load(f)

                # 4. Consolidar todos os dados
                # Atualiza dados_email com informações do anexo (XML/PDF)
                dados_email.update(json_data)
                
                # Garante que os dados do corpo do e-mail (que são prioritários) sobrescrevam os do anexo se houver conflito
                dados_email.update(extract_values(body))

                nmr_nota = dados_email.get('num_nota', {}).get('numero_nota', 'N/A')
                valor_total = str(dados_email.get('valor_total', [{}])[0].get('valor_total', '0')).replace(".", ",")
                
                # Calcula os centros de custo
                dados_email['dados_centros_de_custo'] = process_cost_centers(
                    dados_email.get('cc'),
                    float(valor_total.replace(",", "."))
                )

                # 5. Iniciar a Automação GUI
                logging.info(f"Iniciando automação para a nota fiscal: {nmr_nota}")
                sistema_nf = SystemNF()
                sistema_nf.automation_gui(dados_email=dados_email) # Passa o dicionário completo

                # 6. Finalização e Notificação
                attachment_processed = True
                processed_emails.append(email_id)
                save_processed_emails(processed_emails)
                send_success_message(dani, sender, nmr_nota)
                logging.info(f"Processo para a nota {nmr_nota} concluído com sucesso.")
                break 

        if not attachment_processed:
            raise ValueError("Nenhum anexo válido (PDF/XML) foi encontrado no e-mail.")

    except Exception as e:
        logging.error(f"ERRO FATAL no processamento do e-mail de FRETE: {e}")
        nmr_nota_erro = dados_email.get('num_nota', {}).get('numero_nota', 'N/A')
        if sender:
            send_email_error(dani, sender, str(e), nmr_nota_erro)
        # Relança a exceção para que o dispatcher saiba que o processamento falhou
        raise

def clean_extracted_json(json_data):
    if not isinstance(json_data, dict):
        logging.error(f"Erro: clean_extracted_json recebeu {type(json_data)} em vez de dict.")
        return {}

    if "ISS Retido" in json_data and isinstance(json_data["ISS Retido"], str):
        if json_data["ISS Retido"].lower() == "não":
            json_data["ISS Retido"] = "Não"
        else:
            try:
                json_data["ISS Retido"] = f"{float(json_data['ISS Retido']):.2f}"
            except ValueError:
                logging.warning(f"Valor inválido para ISS Retido: {json_data['ISS Retido']}")

    return json_data

def process_pdf(pdf_path, dados_email):
    logging.info(f"Iniciando processamento do PDF: {pdf_path}")

    if not isinstance(pdf_path, str) or not pdf_path.lower().endswith(".pdf"):
        logging.error(f"Erro: Caminho inválido para PDF: {pdf_path}")
        return None
    if not os.path.exists(pdf_path):
        logging.error(f"Erro: Arquivo não encontrado: {pdf_path}")
        return None

    json_folder = os.path.abspath("notas_json")
    os.makedirs(json_folder, exist_ok=True)

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        logging.error("Erro: API key do Gemini não encontrada.")
        return None

    gemini_api = GeminiAPI(api_key)

    try:
        upload_response = gemini_api.upload_pdf(pdf_path)
        if not upload_response.get("success"):
            logging.error(f"Erro no upload do PDF: {pdf_path}")
            return None

        file_id = upload_response["file_id"]
        status_response = gemini_api.check_processing_status(file_id)
        if status_response.get("state") != "ACTIVE":
            logging.error(f"Erro no processamento do PDF: {pdf_path}")
            return None

        try:
            extracted_text = gemini_api.extract_info(file_id)
            if not extracted_text:
                logging.error(f"Erro ao extrair informações do PDF: {pdf_path}")
                return None

            if "valor_total" not in extracted_text or not extracted_text["valor_total"]:
                logging.warning("Campo 'valor_total' ausente ou vazio no JSON extraído.")
                extracted_text["valor_total"] = "0.00"

            logging.info(f"Conteúdo de 'valor_total' após extração: {extracted_text['valor_total']}" )

            if isinstance(extracted_text, str):
                extracted_text = extracted_text.strip().lstrip("```json").rstrip("```").strip()
                extracted_json = json.loads(extracted_text)
            elif isinstance(extracted_text, dict):
                extracted_json = extracted_text
            else:
                logging.error("Erro: Formato inesperado para o texto extraído.")
                return None

            if not extracted_json:
                logging.error("Erro: JSON extraído está vazio.")
                return None
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao decodificar JSON extraído: {e}")
            return None
        except Exception as e:
            logging.error(f"Erro inesperado ao processar JSON: {e}")
            return None

        cleaned_json = clean_extracted_json(extracted_json)

        valores_extraidos = dados_email.get("valores_extraidos", {})
        data_venc = valores_extraidos.get("data_vencimento")

        json_result = {
            "emitente": cleaned_json.get("emitente", {}),
            "destinatario": cleaned_json.get("destinatario", {}),
            "chave_acesso": {
                "chave": cleaned_json.get("chave_acesso", "")
            },
            "num_nota": {
                "numero_nota": cleaned_json.get("num_nota", "")
            },
            "serie": cleaned_json.get("serie", ""),
            "data_emi": {
                "data_emissao": cleaned_json.get("data_emissao", "")
            },
            "data_venc": {
                "data_venc": data_venc
            },
            "modelo": {
                "modelo": cleaned_json.get("modelo", f"{dados_email.get('modelo')}")
            },
            "tipo_documento": {
                "tipo_documento": cleaned_json.get("tipo_documento", f"{dados_email.get('tipo_documento')}")
            },
            "modelo_email": {
                "modelo_email": cleaned_json.get("modelo_email", f"{dados_email.get('modelo_email')}")
            },
            "valor_total": [
                {
                    "valor_total": cleaned_json.get("valor_total", "0.00")
                }
            ],
            "impostos": {
                "ISS_retido": cleaned_json.get("ISS retido", "0.00"),
                "PIS": cleaned_json.get("PIS", "0.00"),
                "COFINS": cleaned_json.get("COFINS", "0.00"),
                "INSS": cleaned_json.get("INSS", "0.00"),
                "IR": cleaned_json.get("IR", "0.00"),
                "CSLL": cleaned_json.get("CSLL", "0.00"),
            },
            "pagamento_parcelado": cleaned_json.get("pagamento_parcelado", []),
            "info_adicional": {}
        }

        json_filename = re.sub(r'[<>:"/\\|?*]', '', os.path.splitext(os.path.basename(pdf_path))[0]) + ".json"
        json_path = os.path.join(json_folder, json_filename)

        if not os.path.exists(json_folder):
            os.makedirs(json_folder, exist_ok=True)

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_result, f, ensure_ascii=False, indent=4)

        logging.info(f"JSON salvo em: {json_path}")
        return {"json_path": json_path}

    except Exception as e:
        logging.error(f"Erro inesperado ao processar PDF {pdf_path}: {e}")
        return None


def save_attachment(part, directory, dados_email):
    filename = decode_header_value(part.get_filename())
    if not filename:
        logging.error("Nome do arquivo anexado não foi encontrado.")
        dados_email["tipo_arquivo"] = "DESCONHECIDO"  # Define um valor padrão
        return None

    # Determina o tipo de arquivo (PDF)
    if filename.lower().endswith(".pdf"):
        tipo_arquivo = "PDF"
    else:
        tipo_arquivo = "DESCONHECIDO"  # Define um valor padrão para outros tipos de arquivo
        logging.error(f"Arquivo anexado não é um XML ou PDF: {filename}")

    # Armazena o tipo de arquivo no dicionário dados_email
    dados_email["tipo_arquivo"] = tipo_arquivo
    logging.info(f"Tipo de arquivo: {tipo_arquivo}")

    if not os.path.exists(directory):
        os.makedirs(directory)

    filepath = os.path.join(directory, filename)
    logging.info(f"Salvando anexo em: {filepath}")

    with open(filepath, "wb") as f:
        f.write(part.get_payload(decode=True))

    logging.info(f"Anexo salvo em: {filepath}")

    if tipo_arquivo == "XML":
        dados_nota_fiscal = parse_nota_fiscal(filepath)

        if not dados_nota_fiscal:
            logging.error(f"Erro ao processar XML. Nenhum dado extraído.")
            return None

        json_folder = os.path.abspath("notas_json")
        os.makedirs(json_folder, exist_ok=True)

        json_filename = os.path.splitext(os.path.basename(filepath))[0] + ".json"
        json_path = os.path.join(json_folder, json_filename)
        
        logging.info(f"Dados nota fiscal: {dados_nota_fiscal}")

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(dados_nota_fiscal, f, ensure_ascii=False, indent=4)

        logging.info(f"JSON do XML salvo com sucesso: {json_path}")
        return {"json_path": json_path}

    elif tipo_arquivo == "PDF":
        senha_arquivo = dados_email.get("senha_arquivo")
        logging.info(f"Senha recebida para desbloqueio do PDF: {senha_arquivo}")
        if senha_arquivo:
            logging.info(f"Tentando desbloquear arquivo PDF com senha: {senha_arquivo}")  
            try:
                from PyPDF2 import PdfReader, PdfWriter
                
                reader = PdfReader(filepath)
                if reader.is_encrypted:
                    reader.decrypt(senha_arquivo)
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    
                    unlocked_filepath = os.path.join(directory, "unlocked_" + filename)
                    with open(unlocked_filepath, "wb") as f:
                        writer.write(f)
                    
                    logging.info(f"Arquivo PDF desbloqueado e salvo como: {unlocked_filepath}")
                    filepath = unlocked_filepath
                else:
                    logging.info("Arquivo PDF não está criptografado.")
            except Exception as e:
                logging.info(f"Erro ao desbloquear o PDF: {e}")
                return None
        
        json_result = process_pdf(filepath, dados_email)
        if not json_result or "json_path" not in json_result:
            logging.error("Erro: JSON do PDF não foi gerado corretamente.")
            return None
        else:
            logging.info(f"JSON gerado com sucesso: {json_result}")
        return json_result
    else:
        logging.error(f"Tipo de arquivo não suportado: {tipo_arquivo}")
        return None

def process_cost_centers(cc_texto, valor_total):
    valor_total = float(valor_total)
    centros_de_custo = []
    total_calculado = 0.0

    # Verifica se cc_texto é uma string
    if not isinstance(cc_texto, str):
        logging.error(f"Erro: cc_texto deve ser uma string, mas recebeu {type(cc_texto)}.")
        return centros_de_custo

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
            centros_de_custo.append((cc, valor_calculado))
        except ValueError:
            logging.error(f"Erro ao processar centro de custo: {item}")

    diferenca = round(valor_total - total_calculado, 2)

    if abs(diferenca) > 0:
        if centros_de_custo:
            ultimo_cc, ultimo_valor = centros_de_custo[-1]
            if isinstance(ultimo_valor, str):
                ultimo_valor = float(ultimo_valor.replace(",", "."))
            diferenca_float = ultimo_valor + diferenca
            diferenca_str = f"{diferenca_float:.2f}".replace(".", ",")
            centros_de_custo[-1] = (ultimo_cc, diferenca_str)
        else:
            logging.error("Nenhum centro de custo encontrado para aplicar a diferença")
            raise ValueError("Erro: Nenhum centro de custo encontrado para aplicar a diferença")

    return centros_de_custo


def send_email_error(dani, destinatario, erro, nmr_nota):
    config["to"] = destinatario
    dani = Queue(config)

    mensagem = (
        dani.make_message()
        # .set_color("red")  # Remover esta linha
        .add_text(f"Erro durante lançamento de nota fiscal: {nmr_nota}", tag="h1")
        .add_text(str(erro), tag="pre")
    )

    mensagem_assinatura = (
        dani.make_message()
        # .set_color("green")  # Remover esta linha
        .add_text("DANI", tag="h1")
        .add_text("Este é um email enviado automaticamente pelo sistema DANI.", tag="p")
        .add_text(
            "Para reportar erros, envie um email para: ticket.dani@carburgo.com.br",
            tag="p",
        )
        .add_text(
            "Em caso de dúvidas, entre em contato com: caetano.apollo@carburgo.com.br",
            tag="p",
        )
    )
    dani.push(mensagem).push(mensagem_assinatura).flush()


def send_success_message(dani, destinatario, nmr_nota):
    config["to"] = destinatario
    dani = Queue(config)

    mensagem = (
        dani.make_message()
        # .set_color("green")  # Remover esta linha
        .add_text(f"Nota lançada com sucesso número da nota: {nmr_nota}", tag="h1")
        .add_text("Acesse o sistema para verificar o lançamento.", tag="pre")
    )

    mensagem_assinatura = (
        dani.make_message()
        # .set_color("green")  # Remover esta linha
        .add_text("DANI", tag="h1")
        .add_text("Este é um email enviado automaticamente pelo sistema DANI.", tag="p")
        .add_text(
            "Para reportar erros, envie um email para: ticket.dani@carburgo.com.br",
            tag="p",
        )
        .add_text(
            "Em caso de dúvidas, entre em contato com: caetano.apollo@carburgo.com.br",
            tag="p",
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
    def __init__(self, master=None):
        self.br = None

    def automation_gui(self, dados_email=None):
        if not dados_email:
            raise ValueError("Dados do e-mail não foram fornecidos para a automação GUI.")

        # Extração segura de todas as variáveis do dicionário usando .get()
        departamento = dados_email.get('departamento')
        origem = dados_email.get('origem')
        descricao = dados_email.get('descricao')
        cc = dados_email.get('cc')
        cod_item = dados_email.get('cod_item')
        # Acessando valor_total com segurança
        valor_total_list = dados_email.get('valor_total', [{}])
        valor_total = "0.00"
        if valor_total_list and isinstance(valor_total_list, list):
            valor_total = str(valor_total_list[0].get('valor_total', '0.00')).replace(".", ",")

        dados_centros_de_custo = dados_email.get('dados_centros_de_custo', [])
        cnpj_emitente = dados_email.get('emitente', {}).get('cnpj')
        nmr_nota = dados_email.get('num_nota', {}).get('numero_nota')
        data_emi = dados_email.get('data_emi', {}).get('data_emissao')
        data_venc = dados_email.get('data_venc', {}).get('data_venc')
        chave_acesso = dados_email.get('chave_acesso', {}).get('chave')
        modelo = dados_email.get('modelo', {}).get('modelo')
        rateio = dados_email.get('rateio')
        parcelas = dados_email.get('pagamento_parcelado', [])
        serie = dados_email.get('serie')
        data_venc_nfs = dados_email.get('data_vencimento')
        valor_liquido = dados_email.get('valor_liquido', {}).get('valor_liquido')
        tipo_imposto = dados_email.get('tipo_imposto')
        impostos = dados_email.get('impostos', {})
        tipo_documento = dados_email.get('tipo_documento')
        modelo_email = dados_email.get('modelo_email')
        usuario_login = dados_email.get("usuario_login")
        senha_login = dados_email.get("senha_login")
        
        # Validação dos dados necessários
        if not all([departamento, origem, descricao, cc, cod_item, valor_total]):
            raise ValueError("Dados obrigatórios estão faltando")

        gui.PAUSE = 1

        class faker:
            def __call__(self, *args, **kwds):
                return self

            def __getattr__(self, name):
                return self
        
        usuario_login = dados_email.get("usuario_login")
        senha_login = dados_email.get("senha_login")
        
        logging.info(f"Usuário de login: {usuario_login}")
        logging.info(f"Senha de login: {senha_login}")
        
        if usuario_login and senha_login:
            config = {
                "bravos_usr": usuario_login,
                "bravos_pswd": senha_login,
            }
            from bravos.infoBravos import bravos as BravosClass
            self.br = BravosClass(config, m_queue=faker())  
            self.br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")
        else:
            logging.error("Login ou senha do usuário não encontrados.")
            send_email_error(
                dani,
                dados_email.get("email_usuario", "caetano.apollo@carburgo.com.br"),
                "Erro: Login ou senha do usuário não encontrados.",
                nmr_nota,
            )
            return

        logging.info(f"Tipo de Imposto recebido: {tipo_imposto}")
        logging.info(f"Parâmetros recebidos em automation_gui: {locals()}")
        logging.info(f"Tipo de Imposto recebido: {tipo_imposto}")
        if not parcelas:
            logging.info("Nenhuma parcela recebida")
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
        logging.info(f"Data vencimento NFs: {data_venc_nfs}")
        logging.info(f"Tipo de Imposto: {tipo_imposto}")
        logging.info(f"Tipo do documento: {tipo_documento}")
        logging.info(f"Modelo vindo do Email: {modelo_email}")

        cnpj_dest = dados_email.get("destinatario", {}).get("cnpj")

        if cnpj_dest:
            try:
                result = revenda(cnpj_dest) 
                if result:
                    empresa, revenda_nome = result
                    logging.info(f"Empresa: {empresa}, Revenda: {revenda_nome}")
                else:
                    empresa, revenda_nome = "Desconhecido", "Desconhecido"
                    logging.error(f"CNPJ {cnpj_dest} não encontrado no banco de dados.")
            except Exception as e:
                logging.error(f"Erro ao buscar revenda no banco de dados: {e}")
                empresa, revenda_nome = "Erro", "Erro"
        else:
            logging.error("Erro: CNPJ do destinatário não foi fornecido.")
            empresa, revenda_nome = "Sem CNPJ", "Sem CNPJ"
        
        
        if revenda:
            try:
                gui.press("alt")
                gui.press("right")
                gui.press("down")
                gui.press("enter")
                gui.press("down", presses=2)

                gui.write(f"{empresa}.{revenda_nome}")
                gui.press("enter")
                time.sleep(5)
            except NameError as e:
                logging.error(f"Erro: Variável revenda não definida: {e}")
        else:
            logging.error("Erro: Revenda não foi definida corretamente.")
        
        
        class faker:
            def __call__(self, *args, **kwds):
                return self

            def __getattr__(self, name):
                return self
        
        usuario_login = dados_email.get("usuario_login")
        senha_login = dados_email.get("senha_login")
        
        logging.info(f"Tentando login com usuário: {usuario_login}")
        
        if usuario_login and senha_login:
            config_bravos = {
                "bravos_usr": usuario_login,
                "bravos_pswd": senha_login,
            }
            logging.info("Chamando acquire_bravos para abrir o Bravos")
            try:
                self.br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")
                time.sleep(30)
                logging.info("Chamada para abrir o Bravos realizada")
            except Exception as e:
                logging.error(f"Erro ao tentar abrir o Bravos: {e}")
        else:
            logging.error("Login ou senha do usuário não encontrados nos dados do e-mail.")
            send_email_error(
                dani,
                dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                "Erro crítico: Login ou senha não foram passados para a função de automação.",
                nmr_nota,
            )
            return

        
        try:
            data_atual = datetime.now()
            data_formatada = data_atual.strftime("%d%m%Y")
            time.sleep(3)
            window = gw.getWindowsWithTitle("BRAVOS")[0]
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
            gui.press("enter")
            pytesseract.pytesseract_cmd = (
                r"C:\Program Files\Tesseract-OCR/tesseract.exe"
            )
            nova_janela = gw.getActiveWindow()
            janela_left = nova_janela.left
            janela_top = nova_janela.top
            time.sleep(5)
            x, y, width, height = janela_left + 490, janela_top + 300, 120, 21
            screenshot = gui.screenshot(region=(x, y, width, height))
            screenshot = screenshot.convert("L")
            threshold = 190
            screenshot = screenshot.point(lambda p: p > threshold and 255)
            config = r"--psm 7 outputbase digits"
            cliente = pytesseract.image_to_string(screenshot, config=config)
            

            time.sleep(5)
            gui.hotkey("ctrl", "f4")
            gui.press("alt")
            gui.press("a")
            gui.press("down", presses=3)
            gui.press("right")
            gui.press("down", presses=2)
            gui.press("enter")
            time.sleep(5)
            gui.press("tab")
            gui.write(nmr_nota)
            gui.press("tab")
            gui.write(serie)
            gui.press("tab")
            gui.press("down", presses=3)
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
            if origem != "5129":
                send_email_error(
                    dani,
                    dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                    "Erro: Origem da nota errada",
                    nmr_nota,
                )
            else:
                gui.write(origem)
            gui.press("tab", presses=19)
            gui.write(chave_acesso)
            gui.press("tab")
            if dados_email.get("tipo_arquivo") == "PDF":
                if modelo_email != "57":
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        "Erro: Modelo da nota errado",
                        nmr_nota,
                    )
                else:
                    gui.write(modelo_email)
            gui.press("tab", presses=18)

            tipo_arquivo = dados_email.get("tipo_arquivo", "DESCONHECIDO")
            logging.info(f"Tipo de arquivo: {tipo_arquivo}")

            if tipo_arquivo == "PDF":
                gui.press("right", presses=1)
                gui.press("tab", presses=20)
                gui.write(cod_item)
                gui.press("tab", presses=10)
                gui.write("1")
                gui.press("tab")
                gui.write(valor_total)
                gui.press("tab", presses=34)
                gui.write(descricao)
                if tipo_documento == "fatura" or tipo_documento == "boleto":
                    gui.press("tab")
                    gui.press("enter")
                else:
                    if departamento == "7" or departamento == "8" or departamento == "5":
                        gui.press("tab", presses=21)
                        gui.write(valor_total)
                        gui.press("tab")
                        gui.press("enter")

                        gui.press("tab", presses=16)
                        gui.write(valor_total)
                        gui.press("tab", presses=79)
                        gui.press("enter")
                        gui.press("tab")

                gui.press("tab", presses=57)
                gui.press("left")
                gui.press("tab", presses=5)
                gui.press("enter")

                hoje = datetime.now()

                # Determina o próximo mês
                if hoje.month == 12:
                    proximo_mes = 1
                    ano = hoje.year + 1
                else:
                    proximo_mes = hoje.month + 1
                    ano = hoje.year

                # Calcula o primeiro dia do próximo mês
                primeiro_dia_proximo_mes = datetime(ano, proximo_mes, 1)

                # Determina o dia 20 do próximo mês
                dia_20_proximo_mes = datetime(ano, proximo_mes, 20)

                # Calcula a diferença em dias
                dias_restantes = (dia_20_proximo_mes - hoje).days

                gui.press("tab")
                gui.press("enter")
                gui.press("tab")
                gui.write("enter")
                gui.press("tab", presses=10)
                gui.write(str(data_venc_nfs))
                gui.press("tab")
                gui.press("tab")
                gui.press("tab", presses=2)
                
                if valor_liquido == None:
                    valor_liquido = valor_total
                
                gui.write(valor_liquido)
                gui.press("tab")
                gui.press("enter")
                gui.press("tab", presses=3)
                gui.press("enter")
                gui.press("tab", presses=39)
                gui.press("enter")
        except Exception as e:
            logging.error(f"Erro durante a automação: {e}")