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
from db_connection import revenda, login as db_login, cliente, cod_veiculo
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


logging.info("Iniciando o Programa")

HOST = "mail.carburgo.com.br"  # Servidor POP3 corporativo
PORT = 995  # Porta segura (SSL)
USERNAME = "dani@carburgo.com.br"
PASSWORD = "p@r!sA1856"
# Assunto alvo para busca
if "ASSUNTO_ALVO" not in globals():
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

def retrive_user_login(sender):
    resultado = db_login(sender)
    if resultado:
        return resultado
    logging.error(f"Usuário {sender} não encontrado no banco de dados.")
    logging.info(f"resultado da consulta: {resultado}")

# Extrai valores do corpo do e-mail
def extract_values(text):
    if not isinstance(text, str):
        logging.error(f"Erro: extract_values recebeu {type(text)} em vez de string.")
        return {}

    values = {}
    lines = text.splitlines()
    for line in lines:
        line = line.strip()
        low = line.lower()
        if low.startswith("departamento:"):
            values["departamento"] = line.split(":", 1)[1].strip()
        elif low.startswith("origem:"):
            values["origem"] = line.split(":", 1)[1].strip()
        elif low.startswith("descrição:") or low.startswith("descriçao:") or low.startswith("descri��o:"):
            values["descricao"] = line.split(":", 1)[1].strip()
        elif low.startswith("cc:"):
            values["cc"] = line.split(":", 1)[1].strip()
        elif low.startswith("rateio:"):
            values["rateio"] = line.split(":", 1)[1].strip()
        elif low.startswith("código de tributação:") or low.startswith("codigo de tributacao:") or low.startswith("c�digo de tributa��o:"):
            values["cod_item"] = line.split(":", 1)[1].strip()
        elif "data venc" in low or "data de venc" in low or "vencimento" in low:
            parts = re.split(r":|-", line, maxsplit=1)
            raw = parts[1].strip() if len(parts) > 1 else line
            m = re.search(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", raw)
            if m:
                raw_date = m.group(1)
            else:
                raw_date = re.sub(r"\D", "", raw)
            digits = re.sub(r"\D", "", raw_date)
            if len(digits) == 6:
                digits = digits[:4] + "20" + digits[4:]
            values["data_vencimento"] = digits 
        elif low.startswith("tipo imposto:"):
            values["tipo_imposto"] = line.split(":", 1)[1].strip()
        elif low.startswith("senha arquivo:"):
            values["senha_arquivo"] = line.split(":", 1)[1].strip()
        elif low.startswith("modelo:"):
            values["modelo_email"] = line.split(":", 1)[1].strip()
        elif low.startswith("tipo documento:"):
            values["tipo_documento"] = line.split(":", 1)[1].strip()
        elif low.startswith("senha usuario:") or low.startswith("senha usuário:"):
            values["senha_user"] = line.split(":", 1)[1].strip()
        elif low.startswith("chassi:") or low.startswith("chassi do veículo:"):
            values["chassi"] = line.split(":", 1)[1].strip()
        elif low.startswith("despesa:") or low.startswith("tipo despesa:"):
            values["despesa"] = line.split(":", 1)[1].strip()
        
        #! Adicionar as informações adicionais para a parte de outras despesas
        # todo: poderia colocar o campo de observação? ou uso o mesmo da descrição?

    if "data_vencimento" not in values:
        m = re.search(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", text)
        if m:
            digits = re.sub(r"\D", "", m.group(1))
            if len(digits) == 6:
                digits = digits[:4] + "20" + digits[4:]
            values["data_vencimento"] = digits

    logging.info(f"Valores extraídos (linha a linha): {values}")
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
def processar_email_despesa_veiculo(email_message):
    logging.info("="*50)
    logging.info("INICIANDO PROCESSAMENTO DE E-MAIL DE DESPESA DE VEÍCULO")
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
            logging.warning(f"E-mail (ID: {email_id}) já foi processado. Ignorando.")
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
                dados_email.update(json_data)
                dados_email.update(extract_values(body))

                nmr_nota = dados_email.get('num_nota', {}).get('numero_nota', 'N/A')
                valor_total = str(dados_email.get('valor_total', [{}])[0].get('valor_total', '0')).replace(".", ",")
                origem = dados_email.get('origem')

                dados_email['dados_centros_de_custo'] = process_cost_centers(
                    dados_email.get('cc'),
                    float(valor_total.replace(",", ".")),
                    origem # Passa a origem padrão para a função
                )

                # 5. Iniciar a Automação GUI
                logging.info(f"Iniciando automação para a nota fiscal: {nmr_nota}")
                sistema_nf = SystemNF()
                sistema_nf.automation_gui(dados_email=dados_email)

                # 6. Finalização e Notificação
                attachment_processed = True
                processed_emails.append(email_id)
                save_processed_emails(processed_emails)
                send_success_message(dani, sender, nmr_nota, dados_email)
                logging.info(f"Processo para a nota {nmr_nota} concluído com sucesso.")
                break

        if not attachment_processed:
            raise ValueError("Nenhum anexo válido (PDF/XML) foi encontrado no e-mail.")

    except Exception as e:
        logging.error(f"ERRO FATAL no processamento do e-mail de NOTA FISCAL: {e}")
        nmr_nota_erro = dados_email.get('num_nota', {}).get('numero_nota', 'N/A')
        if sender:
            send_email_error(dani, sender, str(e), nmr_nota_erro, dados_email)
        raise

def clean_extracted_json(json_data):
    import logging

    if not isinstance(json_data, dict):
        logging.error(f"Erro: clean_extracted_json recebeu {type(json_data)} em vez de dict.")
        return {}

    # logging.info(f"Conteúdo de json_data antes da limpeza: {json_data}")

    if "ISS_retido" in json_data and isinstance(json_data["ISS_retido"], str):
        if json_data["ISS_retido"].lower() in ["não", "nao"]:
            json_data["ISS_retido"] = "Não"
        else:
            try:
                json_data["ISS_retido"] = f"{float(json_data['ISS_retido']):.2f}"
            except ValueError:
                logging.warning(f"Valor inválido para ISS_retido: {json_data['ISS_retido']}")

    return json_data

def save_attachment(part, directory, dados_email):
    filename = decode_header_value(part.get_filename())
    if not filename:
        logging.error("Nome do arquivo anexado não foi encontrado.")
        dados_email["tipo_arquivo"] = "DESCONHECIDO"
        return None

    if filename.lower().endswith(".pdf"):
        tipo_arquivo = "PDF"
    elif filename.lower().endswith(".xml"):
        tipo_arquivo = "XML"
    else:
        logging.error(f"Arquivo anexado '{filename}' não é um tipo obrigatório (PDF ou XML).")
        dados_email["tipo_arquivo"] = "DESCONHECIDO"
        return None

    dados_email["tipo_arquivo"] = tipo_arquivo
    logging.info(f"Tipo de arquivo obrigatório encontrado: {tipo_arquivo}")

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
        
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(dados_nota_fiscal, f, ensure_ascii=False, indent=4)
        
        logging.info(f"JSON do XML salvo com sucesso: {json_path}")
        return {"json_path": json_path}
            
    return None # Retorno de segurança

def process_cost_centers(cc_texto, valor_total, origem_padrao):
    """
    Processa a string de centros de custo, lidando com todos os cenários:
    1. Único CC (ex: "9")
    2. Múltiplos CCs com valor/percentual (ex: "9 - 70.5, 5 - 29.5")
    3. Múltiplos CCs com origens específicas (ex: "9 - 70.5 (5148), 5 - 29.5")
    """
    valor_total = float(valor_total)
    lancamentos_finais = []
    total_calculado = 0.0

    if not isinstance(cc_texto, str) or not cc_texto.strip():
        logging.error(f"Erro: cc_texto deve ser uma string não vazia, mas recebeu {type(cc_texto)}.")
        return []

    # Regra 1: Trata o caso de um único setor informado
    if cc_texto.strip().isdigit():
        logging.info(f"Centro de custo único detectado: {cc_texto}. Atribuindo 100% do valor.")
        # Transforma "9" em "9 - 100%" para ser processado pela lógica padrão
        cc_texto = f"{cc_texto.strip()} - 100%"

    # Expressão Regular para capturar: CC, valor/percentual e a origem (opcional)
    padrao = re.compile(r"(\d+)\s*-\s*([\d.,%]+)(?:\s*\(\s*(\d+)\s*\))?")
    
    # Separa os lançamentos pela vírgula
    itens = cc_texto.split(',')

    for item in itens:
        item = item.strip()
        match = padrao.match(item)

        if not match:
            logging.error(f"Formato inválido para o lançamento de rateio: '{item}'. Ignorando.")
            continue

        cc, valor_str, origem_especifica = match.groups()
        
        # Regra 3: Define a origem correta
        # Usa a origem específica do lançamento (entre parênteses), se houver.
        # Caso contrário, usa a origem padrão do e-mail.
        origem_final = origem_especifica.strip() if origem_especifica else origem_padrao

        try:
            # Regra 2: Processa múltiplos setores com valor ou percentual
            if "%" in valor_str:
                porcentagem = float(valor_str.replace("%", "").replace(",", "."))
                valor_calculado = round((valor_total * porcentagem) / 100, 2)
            else:
                valor_calculado = float(valor_str.replace(",", "."))

            total_calculado += valor_calculado
            lancamentos_finais.append({
                "cc": cc.strip(),
                "valor": valor_calculado,
                "origem": origem_final
            })
        except ValueError as e:
            logging.error(f"Erro ao converter o valor '{valor_str}' no lançamento '{item}': {e}")
        except Exception as e:
            logging.error(f"Erro inesperado ao processar o lançamento '{item}': {e}")

    # Ajuste final para garantir que a soma dos rateios seja exatamente o valor total da nota
    diferenca = round(valor_total - total_calculado, 2)
    if abs(diferenca) > 0.01 and lancamentos_finais: # Usa uma tolerância
        logging.warning(f"Diferença de R$ {diferenca} encontrada no rateio. Ajustando último lançamento.")
        lancamentos_finais[-1]['valor'] += diferenca

    logging.info(f"Processamento de rateio finalizado. Resultado: {lancamentos_finais}")
    
    return lancamentos_finais

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


def process_installments(parcelas):
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

        logging.info("="*20 + " INICIANDO AUTOMACAO GUI (NOTA FISCAL) " + "="*20)
        logging.info(f"Dados para automação: Nota {nmr_nota}, Valor {valor_total}")

        chassi = dados_email.get('chassi')
        cod_veiculo_valor = None
        if chassi:
            cod_veiculo_valor = cod_veiculo(chassi)
            if not cod_veiculo_valor:
                logging.error(f"Código do veículo não encontrado para o chassi {chassi}")
                send_email_error(
                    dani,
                    dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                    f"Código do veículo não encontrado para o chassi {chassi}",
                    nmr_nota,
                    dados_email
                )
                return
            logging.info(f"Código do veículo para chassi {chassi}: {cod_veiculo_valor}")

        gui.PAUSE = 0.5
        
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
            time.sleep(15)
        else:
            logging.error("Login ou senha do usuário não encontrados.")
            send_email_error(
                dani,
                dados_email.get("email_usuario", "caetano.apollo@carburgo.com.br"),
                "Erro: Login ou senha do usuário não encontrados.",
                nmr_nota,
                dados_email
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
        
        
        if 'empresa' in locals():
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
            window = gw.getWindowsWithTitle("BRAVOS")[0]
            if not window:
                raise Exception("Janela do BRAVOS não encontrada")
            
            cliente_cod = cliente(cnpj_emitente)
            if not cliente_cod:
                raise ValueError(f"Cliente com CNPJ {cnpj_emitente} não encontrado no banco de dados.")

            logging.info(f"Cliente encontrado no banco de dados: {cliente_cod}")
            
            gui.press("alt")
            gui.press("f")
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
            gui.press("down")
            gui.press("tab")
            gui.write("0")
            gui.press("tab")
            gui.write(str(cliente_cod))
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

            tipo_arquivo = dados_email.get("tipo_arquivo", "DESCONHECIDO")
            logging.info(f"Tipo de arquivo: {tipo_arquivo}")

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
                return
            gui.write(cod_item)
            gui.press("tab", presses=10)
            gui.write("1")
            gui.press("tab")
            logging.info(f"Valor total: {valor_total}")
            gui.write(valor_total)
            gui.press("tab", presses=26)
            gui.press("capslock")
            gui.write(descricao)
            gui.press("capslock")
            gui.press(["tab", "enter"])
            gui.press("tab", presses=11)
            gui.press("left", presses=2)
            
            '''
            gui.press("tab", presses=39)
            gui.press("enter")
            manutencao = driver.find_element(By.ID, "mat-tab-label-0-2")
            manutencao.click()
            veiculo_input = driver.find_element(By.ID, "Veiculo")
            veiculo_input.send_keys(cod_veiculo_valor)
            despesa_input = driver.find_element(By.ID, "undefined")
            despesa_input.send_keys(despesa)
            documento_input = driver.find_element(By.ID, "Documento")
            documento_input.send_keys(nmr_nota)
            conta_input = driver.find_element(By.ID, "Contador")
            conta_input.send_keys(contador)
            valor_input = driver.find_element(By.ID, "Valor")
            valor_input.send_keys(valor_total.replace(".", ","))
            data_input = driver.find_element(By.ID, "data")
            data_input.send_keys(data_venc_nfs)
            observacao_input = driver.find_element(By.ID, "Observacao")
            observacao_input.send_keys(descricao)
            salvar = driver.find_element(By.ID, "linx-utils-checkbox-field-3")
            salvar.click()
            gui.hotkey("ctrl", "f4")
            '''
            
            gui.press("tab", presses=5)
            gui.press("enter")
            logging.info(f"Processando parcelas: {parcelas}")
            gui.press("tab")
            gui.press("enter")
            process_installments(parcelas)
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
                for i, lancamento in enumerate(dados_centros_de_custo):
                    cc = lancamento['cc']
                    valor = lancamento['valor']
                    origem_lancamento = lancamento['origem']
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
                    gui.write(origem_lancamento)
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
                        gui.press("enter")
                        logging.info("Último centro de custo salvo e encerrado.")
                        gui.press("tab", presses=3)
                    else:
                        gui.press("tab", presses=3)
            gui.press("enter")
            time.sleep(2)
            gui.press("enter")

        except Exception as e:
            send_email_error(
                dani,
                dados_email.get("sender", "caetano.apollo@carburgo.com.br"),
                e,
                nmr_nota_notificacao,
                dados_email
            )
            return