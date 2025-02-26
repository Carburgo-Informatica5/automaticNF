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

from email.utils import parseaddr

from processar_xml import *
from db_connection import *
from DANImail import Queue, WriteTo
from gemini_api import GeminiAPI
from gemini_main import *

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
ASSUNTO_ALVO = "lançamentos notas fiscais DANI"
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
    }

    lines = text.lower().splitlines()
    for line in lines:
        if isinstance(line, str):  # Garante que cada linha é string
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
            elif "data vencimento" in line.lower():
                data_vencimento = line.split(":", 1)[1].strip()
                values["data_vencimento"] = data_vencimento.replace("/", "")
                logging.info(f"Data de vencimento extraída: {values['data_vencimento']}")
            elif line.startswith("tipo imposto:"):
                values["tipo_imposto"] = line.split(":", 1)[1].strip()
    
    return values



PROCESSED_EMAILS_FILE = os.path.join(current_dir, "processed_emails.json")


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


def check_emails(nmr_nota, extract_values):
    sender = None
    dados_email = {}
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
                            dados_email["body"] = body
                        elif part.get("Content-Disposition") is not None:
                            logging.info("Encontrado anexo no e-mail")
                            dados_nota_fiscal = save_attachment(
                                part, DIRECTORY, dados_email
                            )
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
                                    data_vencimento = valores_extraidos[
                                        "data_vencimento"
                                    ]
                                    tipo_imposto = valores_extraidos["tipo_imposto"]

                                    if "json_path" in dados_nota_fiscal:
                                        with open(dados_nota_fiscal["json_path"], "r") as f:
                                            json_data = json.load(f)

                                        dados_nota_fiscal = {
                                            "valor_total": [
                                                {"valor_total": json_data["valor_total"]["valor_total"]}
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
                                            "pagamento_parcelado": [],
                                            "serie": "",
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
                                        "impostos": json_data.get("impostos", {}),  # Adiciona os impostos extraídos
                                        "valor_liquido": json_data.get("valor_liquido", {}),  # Adiciona o valor líquido extraído
                                    }

                                    logging.info(
                                        f"Dados carregados: {dados_nota_fiscal}"
                                    )
                                    logging.info(
                                        f"Parcelas carregadas: {dados_nota_fiscal.get('pagamento_parcelado')}"
                                    )

                                    # Chama a função de automação com os dados extraídos
                                    sistema_nf = SystemNF()
                                    sistema_nf.automation_gui(
                                        departamento,
                                        origem,
                                        descricao,
                                        cc_texto,
                                        cod_item,
                                        valor_total,
                                        dados_centros_de_custo,
                                        dados_nota_fiscal["emitente"]["cnpj"],
                                        dados_nota_fiscal["num_nota"]["numero_nota"],
                                        dados_nota_fiscal["data_emi"]["data_emissao"],
                                        dados_nota_fiscal["data_venc"]["data_venc"],
                                        dados_nota_fiscal["chave_acesso"]["chave"],
                                        dados_nota_fiscal["modelo"]["modelo"],
                                        rateio,
                                        dados_nota_fiscal.get("pagamento_parcelado", []),
                                        dados_nota_fiscal.get("serie", ""),
                                        valores_extraidos["data_vencimento"],
                                        tipo_imposto,
                                        dados_email.get("impostos", {}).get("INSS", "0.00"),  # Passa o valor de INSS
                                        dados_email.get("impostos", {}).get("IR", "0.00"),  # Passa o valor de IR
                                        dados_email.get("valor_liquido", "0.00"),  # Passa o valor líquido
                                    )

                                    # Adiciona o ID do e-mail processado à lista após lançamento bem-sucedido
                                    processed_emails = load_processed_emails()
                                    processed_emails.append(dados_email["email_id"])
                                    save_processed_emails(processed_emails)
                                    send_success_message(
                                        dani,
                                        sender,
                                        nmr_nota,
                                    )

                                except Exception as e:
                                    logging.error(f"Erro ao processar o e-mail: {e}")
                                    send_email_error(
                                        dani,
                                        sender,
                                        f"Erro ao processar o e-mail: {e}",
                                        nmr_nota,
                                    )
                                    server.quit()
                                    return None
                            else:
                                logging.error("Erro ao salvar ou processar o anexo")
                                send_email_error(
                                    dani,
                                    sender,
                                    "Erro ao salvar ou processar o anexo",
                                    nmr_nota,
                                )
                                server.quit()
                                return None
                else:
                    charset = email_message.get_content_charset()
                    body = decode_body(email_message.get_payload(decode=True), charset)
                    dados_email["body"] = body

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



def map_json_fields(json_data, body):
    # Verifica se o body é uma string
    if not isinstance(body, str):
        logging.error(f"Erro: 'body' deveria ser uma string, mas recebeu {type(body)}.")
        return {}

    valores_extraidos = extract_values(body)

    data_vencimento = valores_extraidos["data_vencimento"]

    mapped_data = {
        "emitente": {
            "cnpj": json_data.get("CNPJ do prestador de serviço"),
            "nome": json_data.get("Nome do prestador de serviço"),
        },
        "destinatario": {
            "cnpj": json_data.get("CNPJ do tomador do serviço"),
            "nome": json_data.get("Nome do tomador do serviço"),
        },
        "num_nota": {
            "numero_nota": json_data.get("Numero da nota"),
        },
        "data_venc": {
            "data_venc": data_vencimento
        },
        "data_emi": {
            "data_emissao": json_data.get("Data da emissão"),
        },
        "valor_total": {
            "valor_total": json_data.get("Valor total"),
        },
        "valor_liquido": {
            "valor_liquido": json_data.get("Valor líquido"),
        },
        "modelo": {
            "modelo": "01",
        },
        "serie": {
            "serie": "1",
        },
        "chave_acesso": {
            "chave": "",
        },
        "impostos": {
            "ISS_retido": json_data.get("ISS retido"),
            "PIS": json_data.get("PIS"),
            "COFINS": json_data.get("COFINS"),
            "INSS": json_data.get("INSS"),
            "IR": json_data.get("IR"),
            "CSLL": json_data.get("CSLL"),
        },
    }
    return mapped_data

def process_pdf(pdf_path, dados_email):
    logging.info(f"Iniciando processamento do PDF: {pdf_path}")

    if not isinstance(pdf_path, str) or not pdf_path.lower().endswith(".pdf"):
        logging.error(f"Erro: Caminho inválido para PDF: {pdf_path}")
        return None
    if not os.path.exists(pdf_path):
        logging.error(f"Erro: Arquivo não encontrado: {pdf_path}")
        return None

    json_folder = os.path.abspath("NOTAS EM JSON")
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

        extracted_text = gemini_api.extract_info(file_id)
        if not extracted_text:
            logging.error(f"Erro ao extrair informações do PDF: {pdf_path}")
            return None

        try:
            extracted_text = extracted_text.strip().lstrip("```json").rstrip("```").strip()
            extracted_json = json.loads(extracted_text) if isinstance(extracted_text, str) else extracted_text

            if not extracted_json:
                logging.error("Erro: JSON extraído está vazio.")
                return None
        except json.JSONDecodeError:
            logging.error(f"Erro ao decodificar JSON extraído: {extracted_text}")
            return None

        cleaned_json = clean_extracted_json(extracted_json)

        # Verifica se o body é uma string antes de chamar map_json_fields
        body = dados_email.get("body", "")
        if not isinstance(body, str):
            logging.error(f"Erro: 'body' deveria ser uma string, mas recebeu {type(body)}.")
            return None

        mapped_json = map_json_fields(cleaned_json, body)

        if not mapped_json:
            logging.error("Erro: JSON final está vazio, não será salvo.")
            return None

        json_filename = re.sub(r'[<>:"/\\|?*]', '', os.path.splitext(os.path.basename(pdf_path))[0]) + ".json"
        json_path = os.path.join(json_folder, json_filename)

        if not os.path.exists(json_folder):
            os.makedirs(json_folder, exist_ok=True)

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(mapped_json, f, ensure_ascii=False, indent=4)

        logging.info(f"JSON salvo em: {json_path}")
        return {"json_path": json_path}

    except Exception as e:
        logging.error(f"Erro inesperado ao processar PDF {pdf_path}: {e}")
        return None

def save_attachment(part, directory, dados_email):
    filename = decode_header_value(part.get_filename())
    if not filename or not filename.lower().endswith((".xml", ".pdf")):
        return None

    content_type = part.get_content_type()
    logging.info(f"Tipo de conteúdo do anexo: {content_type}")

    if not os.path.exists(directory):
        os.makedirs(directory)

    filepath = os.path.join(directory, filename)
    logging.info(f"Salvando anexo em: {filepath}")

    with open(filepath, "wb") as f:
        f.write(part.get_payload(decode=True))

    logging.info(f"Anexo salvo em: {filepath}")

    if filename.lower().endswith(".xml"):
        dados_nota_fiscal = parse_nota_fiscal(filepath)
        return dados_nota_fiscal if dados_nota_fiscal else None
    elif filename.lower().endswith(".pdf"):
        json_result = process_pdf(filepath, dados_email)

        if not json_result or "json_path" not in json_result:
            logging.error("Erro: JSON do PDF não foi gerado corretamente.")
            return None  # Evita tentar acessar um arquivo que não existe

        return json_result  # Retorna o caminho do JSON gerado corretamente

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
        data_venc_nfs,
        ISS_retido,
        INSS,
        IR,
        valor_liquido,
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
        logging.info(f"Data vencimento NFs: {data_venc_nfs}")
        logging.info(f"Tipo de Imposto: {tipo_imposto}")

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

            anexo = dados_email.get("anexo", "")

            if anexo.endswith(".pdf"):
                # Se arquivo endswitch (.pdf)
                gui.press("right", presses=1)
                gui.press("tab", presses=20)
                gui.write(cod_item)
                gui.press("tab", presses=10)
                gui.write("1")
                gui.press("tab")
                gui.write(valor_total)
                gui.press("tab", presses=34)
                gui.write(descricao)

                impostos = dados_email.get("impostos", {})

                def preencher_campo(valor, tabs):
                    """Preenche o campo apenas se o valor for diferente de '0.00' e não for None"""
                    if valor and valor != "0.00":
                        gui.press("tab", presses=tabs)
                        gui.write(valor)

                        # Sequência de preenchimento com verificação
                        gui.press("tab", presses=43)
                        gui.write(valor_total)

                        preencher_campo(impostos.get("PIS"), 1)
                        preencher_campo(valor_total, 5)

                        preencher_campo(impostos.get("COFINS"), 1)
                        preencher_campo(valor_total, 5)

                        preencher_campo(impostos.get("CSLL"), 1)
                        preencher_campo(valor_total, 9)

                        preencher_campo(impostos.get("ISS retido"), 7)

                #! Calculo de dias para impostos
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

                gui.press("tab", presses=23)
                gui.press("left")
                gui.press("tab", presses=5)
                gui.press("enter")

                impostos = dados_email.get("impostos", {})

                # Função para somar apenas valores diferentes de "0.00" ou None
                def somar_impostos(*valores):
                    return sum(
                        float(valor) for valor in valores if valor and valor != "0.00"
                    )

                # Calcula a soma de PIS, COFINS e CSLL (se forem diferentes de 0.00)
                PCC = somar_impostos(
                    impostos.get("PIS"), impostos.get("COFINS"), impostos.get("CSLL")
                )

                # Exibe o valor calculado para depuração
                logging.info(f"Valor de PCC calculado: {PCC:.2f}")

                if INSS != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down")
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(INSS)
                    gui.press("tab", "enter")
                if IR != "0.00":
                    gui.press("tab", presses=7)
                    if tipo_imposto == "normal":
                        gui.press("down", presses=2)
                    elif tipo_imposto == "comissão":
                        gui.press("down", presses=7)
                    else:
                        gui.press("down", presses=5)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(IR)
                    gui.press("tab", "enter")
                if PCC != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down", presses=3)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(PCC)
                    gui.press("tab", "enter")
                if ISS_retido != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down", presses=4)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(ISS_retido)
                    gui.press("tab", "enter")
                gui.press("tab", "enter")
                gui.press("tab", presses=5)
                gui.write(data_venc_nfs)
                gui.press("tab", presses=4)
                gui.write(valor_liquido)
                gui.press("tab", presses=3)
                gui.press("enter")
                gui.press("tab", presses=39)
                gui.press("enter")
                pass
            elif anexo.endswith(".xml"):
                # Se arquivo endswitch (.xml)
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
                        nmr_nota,
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
                dani,
                dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                e,
                nmr_nota,
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

                pdf_path = os.path.join(DIRECTORY, "anexos")
                extracted_info = process_pdf(pdf_path, dados_email)
                if os.path.exists(pdf_path):
                    extracted_info = process_pdf(pdf_path, dados_email)
                else:
                    logging.error("Erro ao precessar o PDF")

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
                data_venc_nfs = dados_email.get("data_vencimento")
                tipo_imposto = dados_email.get("tipo_imposto")
                ISS_retido = dados_email.get("impostos", {}).get("ISS Retido")
                INSS = dados_email.get("impostos", {}).get("INSS")
                IR = dados_email.get("impostos", {}).get("IR")                       
                valor_liquido = dados_email.get("valor_liquido")
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

                def verificar_campos_obrigatorios(
                    departamento,
                    origem,
                    descricao,
                    cc,
                    cod_item,
                    valor_total,
                    dados_centros_de_custo,
                    data_venc_nfs,
                    anexo,
                ):
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

                        # Verifica se o anexo é um PDF e se a data de vencimento está preenchida
                        if anexo.endswith(".pdf") and not data_venc_nfs:
                            mensagem_erro += "- Data de Vencimento\n"

                        logging.error(mensagem_erro)
                        send_email_error(
                            dani,
                            dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                            mensagem_erro,
                            nmr_nota,
                        )

                        # Exemplo de chamada da função
                        verificar_campos_obrigatorios(
                            departamento,
                            origem,
                            descricao,
                            cc,
                            cod_item,
                            valor_total,
                            dados_centros_de_custo,
                            data_venc_nfs,
                            anexo,
                        )

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
                            dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
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
                logging.info(f"Dados dos Centros de Custo: {dados_centros_de_custo}")
                logging.info(f"CNPJ Eminente: {cnpj_emitente}")
                logging.info(f"Número da Nota: {nmr_nota}")
                logging.info(f"Data de Emissão: {data_emi}")
                logging.info(f"Data de Vencimento: {data_venc}")
                logging.info(f"Chave de Acesso: {chave_acesso}")
                logging.info(f"Modelo: {modelo}")

                sistema_nf = SystemNF()

                logging.info(
                    f"Parcelas a serem passadas para automation_gui: {parcelas}"
                )

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
                        data_venc_nfs,
                        ISS_retido,
                        INSS,
                        IR,
                        valor_liquido,
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
                dani,
                "caetano.apollo@carburgo.com.br",
                e,
                nmr_nota,
            )
        logging.info("Esperando antes da nova verificação...")
        time.sleep(30)
