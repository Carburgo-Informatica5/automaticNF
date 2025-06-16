# importações necessárias para o funcionamento do programa
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
import yaml

from email.utils import parseaddr

from processar_xml import *
from db_connection import revenda
from DANImail import Queue, WriteTo
from gemini_api import GeminiAPI
from gemini_main import *

# Configuração do logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Configuração do diretório atual
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append("C:/Users/VAS MTZ/Desktop/Caetano Apollo/automaticNF/bravos")

from bravos import openBravos

logging.info("Iniciando o Programa")

HOST = "mail.carburgo.com.br"  # Servidor POP3 corporativo
PORT = 995  # Porta segura (SSL)
USERNAME = "dani@carburgo.com.br"
PASSWORD = "p@r!sA1856"
# Assunto alvo para busca
ASSUNTO_ALVO = "lançamento nota fiscal"
# Diretório para salvar os anexos
DIRECTORY = os.path.join(current_dir, "anexos")

# Função para decodificar o valor do cabeçalho do e-mail
def decode_header_value(header_value):
    decoded_fragments = decode_header(header_value)
    decoded_string = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_string += fragment.decode(encoding or "utf-8")
        else:
            decoded_string += fragment
    return decoded_string

# Função para normalizar o texto
def normalize_text(text):
    if not isinstance(text, str):  # Garante que text seja string
        logging.error(f"Erro: normalize_text recebeu {type(text)} em vez de string.")
        return ""
    
    return "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    ).lower()

# Função para decodificar o corpo do e-mail
def decode_body(payload, charset):
    if charset is None:
        charset = "utf-8"
    try:
        return payload.decode(charset)
    except (UnicodeDecodeError, LookupError):
        return payload.decode("ISO-8859-1", errors="ignore")

# Arquivo de configuração do servidor de e-mail
config_path = os.path.join(current_dir, "config.yaml")
with open(config_path, "r") as file:
    config = yaml.safe_load(file)
dani = Queue(config)

# Extrai valores do corpo do e-mail
def extract_values(text):
    if not isinstance(text, str):  # Evita chamar `.lower()` em um dicionário
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
    }

    lines = text.lower().splitlines()
    for line in lines:
        line = line.strip()
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

    logging.info(f"Valores extraídos: {values}")

    return values

# Adiciona o caminho do arquivo JSON de emails processados
PROCESSED_EMAILS_FILE = os.path.join(current_dir, "processed_emails.json")

# Carrega os emails processados do arquivo JSON
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

# Salva os id´s dos emails processados no arquivo JSON
def save_processed_emails(processed_emails):
    with open(PROCESSED_EMAILS_FILE, "w") as file:
        json.dump(processed_emails, file)

# Função para verificar os e-mails
def check_emails(nmr_nota, extract_values):
    sender = None
    dados_email = {}
    nmr_nota_notificacao = nmr_nota
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
                        logging.info(f"Parte do e-mail com tipo de conteúdo: {content_type}")
                        if content_type == "text/plain":
                            charset = part.get_content_charset()
                            body = decode_body(part.get_payload(decode=True), charset)
                            dados_email["body"] = body
                            logging.info(f"Corpo do e-mail extraído: {dados_email.get('body')}")

                            # Extrair valores do corpo do e-mail antes de processar o anexo
                            valores_extraidos = extract_values(body)
                            nmr_nota = valores_extraidos.get("numero_nota") or valores_extraidos.get("num_nota") or nmr_nota
                            dados_email.update(valores_extraidos)
                            logging.info(f"Dados do e-mail após extração: {dados_email}")
                            if not nmr_nota:
                                logging.error("Número da nota fiscal não encontrado nos dados extraídos.")
                        elif part.get("Content-Disposition") is not None:
                            logging.info("Encontrado anexo no e-mail")
                            # Processar o anexo com os dados atualizados
                            dados_nota_fiscal = save_attachment(part, DIRECTORY, dados_email)
                            if dados_nota_fiscal:
                                try:
                                    valores_extraidos = extract_values(body)
                                    dados_email.update(valores_extraidos)
                                    logging.info(f"Dados do e-mail após extração: {dados_email}")
                                    departamento = valores_extraidos["departamento"]
                                    origem = valores_extraidos["origem"]
                                    descricao = valores_extraidos["descricao"]
                                    cc_texto = valores_extraidos["cc"]
                                    logging.info(
                                        f"Texto de centros de custo extraído: {cc_texto}"
                                    )
                                    rateio = valores_extraidos["rateio"]
                                    cod_item = valores_extraidos["cod_item"]
                                    data_vencimento = valores_extraidos["data_vencimento"]
                                    tipo_imposto = valores_extraidos["tipo_imposto"]
                                    logging.info(f"Tipo de imposto extraído: {valores_extraidos.get('tipo_imposto')}")

                                    if not dados_nota_fiscal:
                                        logging.error("Erro: 'dados_nota_fiscal' está vazio. O anexo pode não ter sido processado corretamente.")
                                    elif "json_path" not in dados_nota_fiscal:
                                        logging.error("Erro: 'dados_nota_fiscal' não contém 'json_path'.")
                                    else:
                                        with open(dados_nota_fiscal["json_path"], "r") as f:
                                            json_data = json.load(f)
                                            logging.info(f"Conteúdo de json_data['valor_total']: {json_data['valor_total']}")
                                            
                                        tipo_arquivo = dados_email.get("tipo_arquivo")
                                        logging.info(f"Tipo de arquivo extraído: {tipo_arquivo}")


                                        dados_nota_fiscal = {
                                            "valor_total": [
                                                {"valor_total": json_data["valor_total"][0]["valor_total"]} if isinstance(json_data["valor_total"], list) else {"valor_total": json_data["valor_total"]["valor_total"]}
                                            ],
                                            "emitente": {
                                                "nome": json_data["emitente"]["nome"],
                                                "cnpj": json_data["emitente"]["cnpj"],
                                            },
                                            "num_nota": {
                                                "numero_nota": json_data["num_nota"]["numero_nota"]
                                            },
                                            "data_emi": {
                                                "data_emissao": json_data["data_emi"]["data_emissao"]
                                            },
                                            "data_venc": {
                                                "data_venc": valores_extraidos["data_vencimento"]
                                            },
                                            "chave_acesso": {
                                                "chave": json_data["chave_acesso"]["chave"]
                                            },
                                            "modelo": {
                                                "modelo": json_data["modelo"]["modelo"]
                                            },
                                            "destinatario": {
                                                "nome": json_data["destinatario"]["nome"],
                                                "cnpj": json_data["destinatario"]["cnpj"],
                                            },
                                            "pagamento_parcelado": None if tipo_arquivo == "pdf" else json_data.get('pagamento_parcelado', None),
                                            "serie": json_data["serie"],
                                        }

                                    if not dados_nota_fiscal["valor_total"]:
                                        raise ValueError(
                                            "Valor total não encontrado na nota fiscal"
                                        )

                                    valor_total = str(
                                        dados_nota_fiscal["valor_total"][0][
                                            "valor_total"
                                        ]
                                    ).replace(".", ",")
                                    dados_centros_de_custo = process_cost_centers(cc_texto, float(valor_total.replace(",", ".")))
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
                                        "email_id": email_id,
                                        "parcelas": dados_nota_fiscal.get("pagamento_parcelado", []),
                                        "serie": dados_nota_fiscal.get("serie", ""),
                                        "data_venc_nfs": valores_extraidos["data_vencimento"],
                                        "tipo_imposto": valores_extraidos["tipo_imposto"],
                                        "impostos": json_data.get("impostos", {}),  
                                        "valor_liquido": json_data.get("valor_liquido", {}),
                                        "tipo_arquivo": dados_email.get("tipo_arquivo", "DESCONHECIDO"),
                                        "tipo_documento": json_data.get("tipo_documento", {}).get("tipo_documento"),
                                        "modelo_email": json_data.get("modelo_email", {}).get("modelo_email"),
                                    }

                                    logging.info(
                                        f"Dados carregados: {dados_nota_fiscal}"
                                    )
                                    logging.info(
                                        f"Parcelas carregadas: {dados_nota_fiscal.get('pagamento_parcelado')}"
                                    )

                                    if tipo_imposto is None:
                                        logging.error("Erro: tipo_imposto não foi encontrado nos dados extraídos!")
                                    logging.info(f"Conteúdo de dados_email antes de automation_gui: {dados_email}")
                                    logging.info(f"Tipo de imposto antes de chamar automation_gui: {tipo_imposto}")
                                    
                                    tipo_imposto = valores_extraidos.get("tipo_imposto", "NÃO INFORMADO")
                                    logging.info(f"Tipo de arquivo a ser processado: {tipo_arquivo}")
                                    
                                    if dados_email is None:
                                        logging.error("Erro: `dados_email` já é None antes de chamar automation_gui")
                                    else:
                                        logging.info(f"`dados_email` antes de chamar automation_gui: {dados_email}")
                                        logging.info(f"Referência de `dados_email`: {id(dados_email)}")
                                        
                                    
                                    sistema_nf = SystemNF()
                                    logging.info(f"Chamando automation_gui com os seguintes parâmetros:")
                                    logging.info(f"departamento: {dados_email.get('departamento')}")
                                    logging.info(f"origem: {dados_email.get('origem')}")
                                    logging.info(f"descricao: {dados_email.get('descricao')}")
                                    logging.info(f"cc: {dados_email.get('cc')}")
                                    logging.info(f"cod_item: {dados_email.get('cod_item')}")
                                    logging.info(f"valor_total: {dados_email.get('valor_total')}")
                                    logging.info(f"dados_centros_de_custo: {dados_email.get('dados_centros_de_custo')}")
                                    logging.info(f"cnpj_emitente: {dados_email.get('emitente', {}).get('cnpj')}")
                                    logging.info(f"nmr_nota: {dados_email.get('num_nota', {}).get('numero_nota')}")
                                    logging.info(f"data_emi: {dados_email.get('data_emi', {}).get('data_emissao')}")
                                    logging.info(f"data_venc: {dados_email.get('data_venc', {}).get('data_venc')}")
                                    logging.info(f"chave_acesso: {dados_email.get('chave_acesso', {}).get('chave')}")
                                    logging.info(f"modelo: {dados_email.get('modelo', {}).get('modelo')}")
                                    logging.info(f"rateio: {dados_email.get('rateio')}")
                                    logging.info(f"parcelas: {dados_email.get('parcelas')}")
                                    logging.info(f"serie: {dados_email.get('serie')}")
                                    logging.info(f"data_venc_nfs: {dados_email.get('data_venc_nfs')}")
                                    logging.info(f"ISS_retido: {dados_email.get('impostos', {}).get('ISS_retido')}")
                                    logging.info(f"INSS: {dados_email.get('impostos', {}).get('INSS')}")
                                    logging.info(f"IR: {dados_email.get('impostos', {}).get('IR')}")
                                    logging.info(f"valor_liquido: {dados_email.get('valor_liquido', {}).get('valor_liquido')}")
                                    logging.info(f"tipo_imposto: {dados_email.get('tipo_imposto')}")
                                    logging.info(f"PIS: {dados_email.get('impostos', {}).get('PIS')}")
                                    logging.info(f"COFINS: {dados_email.get('impostos', {}).get('COFINS')}")
                                    logging.info(f"CSLL: {dados_email.get('impostos', {}).get('CSLL')}")
                                    sistema_nf.automation_gui(
                                        dados_email.get('departamento'),
                                        dados_email.get('origem'),
                                        dados_email.get('descricao'),
                                        dados_email.get('cc'),
                                        dados_email.get('cod_item'),
                                        dados_email.get('valor_total'),
                                        dados_email.get('dados_centros_de_custo'),
                                        dados_email.get('emitente', {}).get('cnpj'),
                                        dados_email.get('num_nota', {}).get('numero_nota'),
                                        dados_email.get('data_emi', {}).get('data_emissao'),
                                        dados_email.get('data_venc', {}).get('data_venc'),
                                        dados_email.get('chave_acesso', {}).get('chave'),
                                        dados_email.get('modelo', {}).get('modelo'),
                                        dados_email.get('rateio'),
                                        dados_email.get('parcelas'),
                                        dados_email.get('serie'),
                                        dados_email.get('data_venc_nfs'),
                                        dados_email.get('valor_liquido', {}).get('valor_liquido'),
                                        dados_email.get('tipo_imposto'),
                                        dados_email.get('impostos', {}),
                                        dados_email.get('tipo_documento'),
                                        dados_email.get('modelo_email'),
                                        dados_email
                                    )

                                    # Adiciona o ID do e-mail processado à lista após lançamento bem-sucedido
                                    processed_emails = load_processed_emails()
                                    processed_emails.append(dados_email["email_id"])
                                    save_processed_emails(processed_emails)
                                    send_success_message(
                                        dani,
                                        sender,
                                        nmr_nota_notificacao,
                                        dados_email
                                    )

                                except Exception as e:
                                    logging.error(f"Erro ao processar o e-mail: {e}")
                                    
                                    nmr_nota_notificacao = None
                                    
                                    if 'nmr_nota_notificacao' in locals() and nmr_nota:
                                        nmr_nota_notificacao = nmr_nota
                                    elif dados_email and dados_email.get('num_nota') and dados_email['num_nota'].get('numero_nota'):
                                        nmr_nota_notificacao = dados_email['num_nota']['numero_nota']
                                    else:
                                        nmr_nota_notificacao = "NÃO INFORMADO"
                                    
                                    send_email_error(
                                        dani,
                                        sender,
                                        f"Erro ao processar o e-mail: {e}",
                                        nmr_nota_notificacao,
                                        dados_email
                                    )
                                    return None
                            else:
                                logging.error("Erro ao salvar ou processar o anexo")
                                send_email_error(
                                    dani,
                                    sender,
                                    "Erro ao salvar ou processar o anexo",
                                    nmr_nota_notificacao,
                                    dados_email
                                )
                                return None
                else:
                    charset = email_message.get_content_charset()
                    body = decode_body(email_message.get_payload(decode=True), charset)
                    dados_email["body"] = body

        if not dados_email:
            logging.error("Erro: Nenhum dado extraído do e-mail.")
            return None
    except Exception as e:
        logging.error(f"Erro ao verificar emails: {e}")
        if "-ERR EOF" in str(e):
            logging.warning("Erro de desconexão do servidor POP3 detectado, não será enviado e-mail de erro ao usuário.")
        else:
            send_email_error(
                dani,
                sender if sender else "caetano.apollo@carburgo.com.br",
                f"Erro ao verificar emails: {e}",
                nmr_nota_notificacao,
                dados_email
            )
        return None
    finally:
        try:
            if server:
                server.quit()
                logging.info("Conexão com o servidor POP3 encerrada.")
        except Exception as e:
            logging.error(f"Erro ao encerrar conexão POP3: {e}")

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
            "serie": cleaned_json.get("serie", "0"),
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

    # Determina o tipo de arquivo (PDF ou XML)
    if filename.lower().endswith(".pdf"):
        tipo_arquivo = "PDF"
    elif filename.lower().endswith(".xml"):
        tipo_arquivo = "XML"
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
            partes = [x.strip() for x in item.split("-")]
            if len(partes) != 2:
                logging.error(f"Formato inválido para centro de custo: {item}")
                continue
            cc, valor = partes
            logging.info(f"Processando centro de custo: {cc}, valor: {valor}")
            if "%" in valor:
                porcentagem = float(valor.replace("%", ""))
                valor_calculado = round((valor_total * porcentagem) / 100, 2)
            else:
                valor_calculado = float(valor.replace(",", "."))
            total_calculado += valor_calculado
            centros_de_custo.append((cc, valor_calculado))
        except Exception as e:
            logging.error(f"Erro ao processar centro de custo: {item} ({e})")

    diferenca = round(valor_total - total_calculado, 2)

    if abs(diferenca) > 0:
        if centros_de_custo:
            ultimo_cc, ultimo_valor = centros_de_custo[-1]
            if isinstance(ultimo_valor, str):
                ultimo_valor = float(ultimo_valor.replace(",", "."))
            diferenca_float = ultimo_valor + diferenca
            centros_de_custo[-1] = (ultimo_cc, diferenca_float)
        else:
            logging.error("Nenhum centro de custo encontrado para aplicar a diferença")
            raise ValueError("Erro: Nenhum centro de custo encontrado para aplicar a diferença")

    if isinstance(centros_de_custo, list):
        for item in centros_de_custo:
            if not (isinstance(item, tuple) and len(item) == 2):
                logging.error(f"Item inválido em centros_de_custo: {item}")
                return []
    else:
        logging.error(f"centros_de_custo não é uma lista: {centros_de_custo}")
        return []

    return centros_de_custo

def send_email_error(dani, destinatario, erro, nmr_nota_notificacao, dados_email=None):
    if not nmr_nota_notificacao and dados_email:
        nmr_nota_notificacao = (
            dados_email.get("num_nota", {}).get("numero_nota")
            or dados_email.get("nmr_nota")
            or "Não foi possível recuperar o número da nota"
        )
        
    config["to"] = destinatario
    dani = Queue(config)
    logging.info(f"Enviando mensagem de erro para {destinatario} com número da nota: {nmr_nota_notificacao}")

    mensagem = (
        dani.make_message()
        .set_status("error")
        .add_text(f"Erro durante lançamento de nota fiscal", tag="h1")
        .add_text(f"Número da nota: {nmr_nota_notificacao}", tag="pre")
        .add_text(f"Detalhes do erro: {str(erro)}", tag="pre")
    )

    assinatura_dani = (
        "Este é um email enviado automaticamente pelo sistema DANI.<br>"
        "Para reportar erros, envie um email para: ticket.dani@carburgo.com.br<br>"
        "Em caso de dúvidas, entre em contato com: caetano.apollo@carburgo.com.br"
    )
    
    mensagem.add_text(assinatura_dani, tag="p")

    try:
        dani.push(mensagem).flush()
    except Exception as e:
        logging.error(f"Erro ao processar o e-mail: {e}")
        logging.error("Erro crítico: Não foi possível enviar e-mail de erro.")


def send_success_message(dani, destinatario, nmr_nota_notificacao, dados_email=None):
    if not nmr_nota_notificacao and dados_email:
        nmr_nota_notificacao = (
            dados_email.get("num_nota", {}).get("numero_nota")
            or dados_email.get("nmr_nota")
            or "Não foi possível recuperar o número da nota"
        )

    config["to"] = destinatario
    dani = Queue(config)
    logging.info(f"Enviando mensagem de sucesso para {destinatario} com número da nota: {nmr_nota_notificacao}")

    texto_sucesso = (
        f"Número da nota: {nmr_nota_notificacao}\n"
        "Acesse o sistema para verificar o lançamento."
    )

    mensagem = (
        dani.make_message()
        .set_status("success")
        .add_text(f"Nota lançada com sucesso", tag="h1")
        .add_text(texto_sucesso, tag="pre")
    )

    assinatura_dani = (
        "Este é um email enviado automaticamente pelo sistema DANI.<br>"
        "Para reportar erros, envie um email para: ticket.dani@carburgo.com.br<br>"
        "Em caso de dúvidas, entre em contato com: caetano.apollo@carburgo.com.br"
    )
    
    mensagem.add_text(assinatura_dani, tag="p")

    try:
        dani.push(mensagem).flush()
    except Exception as e:
        logging.error(f"Erro ao enviar e-mail de sucesso: {e}")


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

    def automation_gui(
        self, departamento, origem, descricao, cc, cod_item, valor_total,
        dados_centros_de_custo, cnpj_emitente, nmr_nota, data_emi, data_venc,
        chave_acesso, modelo, rateio, parcelas, serie, data_venc_nfs,
        valor_liquido, tipo_imposto, impostos, tipo_documento, modelo_email, dados_email=None
    ):

        gui.PAUSE = 1

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
        nmr_nota_notificacao = nmr_nota if nmr_nota else dados_email.get('num_nota', {}).get('numero_nota', 'NÃO INFORMADO')

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

            time.sleep(5)
            gui.hotkey("ctrl", "f4")
            gui.press("alt")
            if revenda_nome == 4:
                gui.press("right", presses=4)
            else:
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
            if dados_email.get("tipo_arquivo") == "PDF":
                gui.write(modelo_email)
            else:
                gui.write(modelo)
            gui.press("tab", presses=18)

            tipo_arquivo = dados_email.get("tipo_arquivo", "DESCONHECIDO")
            logging.info(f"Tipo de arquivo: {tipo_arquivo}")

            if tipo_arquivo == "PDF":
                if tipo_documento == "fatura" or tipo_documento == "boleto":
                    gui.press("right", presses=2)
                    gui.press("tab", presses=5)
                    gui.press("enter")
                    gui.press("tab", presses=4)
                else:
                    gui.press("right", presses=1)
                    gui.press("tab", presses=20)
                gui.write(cod_item)
                gui.press("tab", presses=10)
                gui.write("1")
                gui.press("tab")
                gui.write(valor_total)
                if tipo_documento == "fatura" or tipo_documento == "boleto":
                    gui.press("tab", presses=26)
                else:
                    gui.press("tab", presses=34)
                gui.write(descricao)
                if tipo_documento == "fatura" or tipo_documento == "boleto":
                    gui.press("tab")
                    gui.press("enter")
                else:
                    if impostos.get("IR") != "0.00":
                        gui.press("tab", presses=17)
                        gui.write(valor_total)
                        gui.press("tab", presses=2)
                        gui.write(impostos.get("IR"))

                    elif impostos.get("PIS") != "0.00":
                        gui.press("tab", presses=24)
                        gui.write(valor_total)
                        gui.press("tab")
                        gui.write(impostos.get("PIS"))

                    elif impostos.get("COFINS") != "0.00":
                        gui.press("tab", presses=5)
                        gui.write(valor_total)
                        gui.press("tab")
                        gui.write(impostos.get("COFINS"))

                    elif impostos.get("CSLL") != "0.00":
                        gui.press("tab", presses=5)
                        gui.write(valor_total)
                        gui.press("tab")
                        gui.write(impostos.get("CSLL"))

                        gui.press("tab", presses=40)
                        gui.press("left")
                gui.press("tab", presses=11)
                gui.press("left", presses=2)
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


                impostos = dados_email.get("impostos", {})

                # Função para somar apenas valores diferentes de "0.00" ou None
                def somar_impostos(*valores):
                    return sum(
                        float(valor) for valor in valores if valor and valor != "0.00"
                    )

                PCC = somar_impostos(
                    impostos.get("PIS"), impostos.get("COFINS"), impostos.get("CSLL")
                )

                # Exibe o valor calculado para depuração
                logging.info(f"Valor de PCC calculado: {PCC:.2f}")

                if impostos.get("IR") != "0.00":
                    gui.press("tab", presses=7)
                    if tipo_imposto == "normal":
                        gui.press("down", presses=2)
                        logging.info("Tipo de imposto normal")
                    elif tipo_imposto == "comissão":
                        gui.press("down", presses=7)
                        logging.info("Tipo de imposto comissão")
                    else:
                        gui.press("down", presses=5)
                        logging.info("Tipo de imposto não especificado")
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(impostos.get("IR"))
                    logging.info("Passou do if do IR")
                    gui.press("tab", "enter")
                if PCC != 0.00:
                    gui.press("tab", presses=7)
                    gui.press("down", presses=3)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(PCC)
                    gui.press("tab", "enter")
                if impostos.get("ISS_retido") != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down", presses=4)
                    gui.press("tab")
                    gui.write(dias_restantes - 5)
                    gui.press("tab", presses=2)
                    gui.write(impostos.get("ISS_retido"))
                    gui.press("tab", "enter")
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


            elif tipo_arquivo == "XML":
                gui.press("right", presses=2)
                gui.press("tab", presses=5)
                gui.press("enter")
                gui.press("tab", presses=4)
                if (
                    modelo == "55" or modelo == "43" or modelo == "22"
                ) and cod_item != "7":
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        "Erro: Modelo de nota fiscal inválido para o código de tributação informado.",
                        nmr_nota_notificacao,
                        dados_email
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
                gui.press("tab", presses=36)
                if rateio.lower() == "sim":
                    logging.info(f"dados_centros_de_custo recebido para rateio: {dados_centros_de_custo} ({type(dados_centros_de_custo)})")
                    if not isinstance(dados_centros_de_custo, list):
                        logging.error(f"dados_centros_de_custo não é uma lista: {dados_centros_de_custo}")
                        send_email_error(
                            dani,
                            dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                            "Erro: dados_centros_de_custo não é uma lista",
                            nmr_nota_notificacao,
                            dados_email
                        )
                        return
                    gui.press("enter")
                    gui.press("tab", presses=9)
                    logging.info("Pressionou o tab corretamente")
                    total_rateio = 0
                    for i, (cc, valor) in enumerate(dados_centros_de_custo):
                        try:
                            valor_float = float(str(valor).replace(",", "."))
                        except Exception as e:
                            logging.error(f"Erro ao converter valor do centro de custo: {valor} ({e})")
                            valor_float = 0.0
                        logging.info("Lançando centro de custo")
                        logging.info(f"Centro de custo: {cc}, Valor: {valor_float}")
                        logging.info(f"Dados centro de custo: {dados_centros_de_custo}")
                        logging.info(f"Revenda: {revenda}")
                        logging.info(f"Revenda nome: {revenda_nome}")
                        if i >= 1:
                            gui.press("tab")
                        gui.write(str(revenda_nome))
                        logging.info(f"Revenda nome escrito: {revenda_nome}")
                        gui.press("tab", presses=3)
                        logging.info(f"pressionando tab para escrever centro de custo: {cc}")
                        gui.write(cc)
                        logging.info(f"Centro de custo escrito: {cc}")
                        gui.press("tab", presses=2)
                        gui.write(origem)
                        gui.press("tab", presses=2)
                        if i == len(dados_centros_de_custo) - 1:
                            valor_float = float(valor_total.replace(",", ".")) - total_rateio
                        else:
                            total_rateio += valor_float
                        gui.write(f"{valor_float:.2f}".replace(".", ","))
                        gui.press("f2", interval=2)
                        gui.press("f3")
                        if i == len(dados_centros_de_custo) - 1:
                            gui.press("f2", interval=2)
                            gui.press("esc", presses=3)
                            logging.info("Último centro de custo salvo e encerrado.")
                            gui.press("tab", presses=3)
                        else:
                            gui.press("tab", presses=3)
                gui.press("enter")

        except Exception as e:
            send_email_error(
                dani,
                dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                e,
                nmr_nota_notificacao,
                dados_email
            )
            print(f"Erro durante a automação: {e}")
            print("Automação iniciada com os dados extraídos.")

if __name__ == "__main__":
    while True:
        logging.info("Iniciando a automação")
        try:
            nmr_nota = ""
            dados_email = check_emails(nmr_nota, extract_values)
            if dados_email is not None:
                logging.info("Executando automação GUI")
                logging.info(f"Conteúdo completo de dados_email: {dados_email}")

                tipo_arquivo = dados_email.get("tipo_arquivo", "DESCONHECIDO")
                logging.info(f"Tipo de arquivo a ser processado: {tipo_arquivo}")
                
                if tipo_arquivo == "DESCONHECIDO":
                    logging.error("Tipo e arquivo desconhecido ou não especificado")
                    nmr_nota_notificacao = nmr_nota if 'nmr_nota' in locals() else None
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        "Erro: Tipo de arquivo desconhecido ou não suportado",
                        nmr_nota_notificacao,
                        dados_email
                    )

                pdf_path = os.path.join(DIRECTORY, "anexos")
                extracted_info = process_pdf(pdf_path, dados_email)
                if os.path.exists(pdf_path):
                    extracted_info = process_pdf(pdf_path, dados_email)
                else:
                    logging.error("Erro ao processar o PDF")

                departamento = dados_email.get("departamento")
                origem = dados_email.get("origem")
                descricao = dados_email.get("descricao")
                cc = dados_email.get("cc") 
                logging.info(f"CC recebido: {cc}")
                cod_item = dados_email.get("cod_item")
                valor_total = dados_email.get("valor_total")
                dados_centros_de_custo = dados_email.get("dados_centros_de_custo")
                rateio = dados_email.get("rateio")
                parcelas = dados_email.get("parcelas", [])
                serie = dados_email.get("serie")
                data_venc_nfs = dados_email.get("data_vencimento")
                tipo_imposto = dados_email.get("tipo_imposto", "normal")
                impostos = dados_email.get("impostos")
                valor_liquido = dados_email.get("valor_liquido", "valor_total")
                if "parcelas" in dados_email:
                    parcelas = dados_email["parcelas"]
                else:
                    logging.warning("Chave 'parcelas' não encontrada em dados_email")
                logging.info(f"Parcelas recebidas: {parcelas}")
                
                def processar_dados(dados_email):
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
                        if "num_nota" in dados_email and "numero_nota" in dados_email["num_nota"]:
                            nmr_nota = dados_email.get("num_nota", {}).get("numero_nota")
                            if not nmr_nota:
                                logging.error("Número da nota não encontrado. Nenhum e-mail será enviado.")
                                return None
                        else:
                            logging.error("Número da nota fiscal não encontrado em dados_email.")
                            return
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
                            nmr_nota_notificacao,
                            dados_email
                        )
        
                # Adicionando logs para verificar os dados extraídos
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
                logging.info(f"Tipo de Imposto: {tipo_imposto}")
                
                logging.info(f"Dados disponíveis antes de chamar automation_gui: {dados_email}")
                logging.info(f"Tipo de imposto extraído: {dados_email.get('tipo_imposto')}")
        
                sistema_nf = SystemNF()
        
                logging.info(
                    f"Parcelas a serem passadas para automation_gui: {parcelas}"
                )
        
                try:
                    logging.info(f"Impostos: {dados_email.get('impostos')}")
                    logging.info(f"Valor Líquido: {valor_liquido}")
                    logging.info(f"Tipo de Imposto: {tipo_imposto}")
                    
        
                    if dados_email is not None:
                        modelo_email = dados_email.get("modelo_email", "DESCONHECIDO")
                        tipo_documento = dados_email.get("tipo_documento", "DESCONHECIDO")
                        sistema_nf.automation_gui(
                            departamento, origem, descricao, cc, cod_item, valor_total,
                            dados_centros_de_custo, cnpj_emitente, nmr_nota, data_emi, data_venc,
                            chave_acesso, modelo, rateio, parcelas, serie, data_venc_nfs,
                            valor_liquido, tipo_imposto, impostos, tipo_documento, modelo_email, dados_email     
                        )
                    else:
                        logging.error("Erro: Dados do e-mail é none. Verifique o processamento do e-mail")
                    # Adicionar o ID do e-mail processado à lista após lançamento bem-sucedido
                    processed_emails = load_processed_emails()
                    processed_emails.append(dados_email["email_id"])
                    save_processed_emails(processed_emails)
                    send_success_message(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        nmr_nota_notificacao,
                        dados_email
                    )
                except Exception as e:
                    logging.error(f"Erro ao chamar automation_gui: {e}")
                    send_email_error(
                        dani,
                        dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                        e,
                        nmr_nota_notificacao,
                        dados_email
                    )
            else:
                logging.info("Nenhum dado extraído, automação não será executada")
                logging.error("Erro: dados_email é None. Verifique o processamento do e-mail.")
        except Exception as e:
            logging.error(f"Erro durante a automação: {e}")
            send_email_error(
                dani,
                "caetano.apollo@carburgo.com.br",
                e,
                nmr_nota_notificacao,
                dados_email
            )
        logging.info("Esperando antes da nova verificação...")
        time.sleep(30)