import poplib
import email
import time
import logging
import json
import os
from email.parser import BytesParser
from email.header import decode_header

# Importa as funções principais dos outros scripts
from main import processar_email_nota_fiscal
from frete import processar_email_frete
from veiculos_despesa import processar_email_despesa_veiculo  
# --- Configurações do Email ---
HOST = "mail.carburgo.com.br"
PORT = 995
USERNAME = "dani@carburgo.com.br"
PASSWORD = "p@r!sA1856"

# --- Assuntos Alvo ---
ASSUNTO_NOTA_FISCAL = "lançamento nota fiscal" or "lançamento nf" or "lançamento nfe" or "lançamento nfs" or "lançamento nfs-e" or "lançamento de nota fiscal"
ASSUNTO_FRETE = "lançamento frete" or "lançamento de frete" or "lancamento frete" or "lancamento de frete"
ASSUNTO_VEICULO = "lançamento despesa veiculo" or "lançamento despesa veículo" or "lancamento despesa veiculo" or "lancamento despesa veículo" or "lançamento despesa de veiculo" or "lançamento despesa de veículo" or "lancamento despesa de veiculo" or "lancamento despesa de veículo"

# --- Arquivo para guardar IDs dos e-mails já processados ---
current_dir = os.path.dirname(os.path.abspath(__file__))
PROCESSED_EMAILS_FILE = os.path.join(current_dir, "processed_emails.json")

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

def decode_header_value(header_value):
    """Decodifica o valor do cabeçalho do e-mail para texto legível."""
    decoded_fragments = decode_header(header_value)
    decoded_string = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_string += fragment.decode(encoding or "utf-8")
        else:
            decoded_string += fragment
    return decoded_string

def load_processed_emails():
    """Carrega a lista de IDs de e-mails já processados do arquivo JSON."""
    if os.path.exists(PROCESSED_EMAILS_FILE):
        with open(PROCESSED_EMAILS_FILE, "r") as file:
            try:
                return json.load(file)
            except json.JSONDecodeError:
                logging.warning("Arquivo processed_emails.json está corrompido ou vazio. Criando um novo.")
                return []
    return []

def save_processed_emails(processed_ids):
    """Salva a lista atualizada de IDs de e-mails no arquivo JSON."""
    with open(PROCESSED_EMAILS_FILE, "w") as file:
        json.dump(processed_ids, file, indent=4)

if __name__ == '__main__':
    while True:
        logging.info("Verificando novos e-mails...")
        processed_ids = load_processed_emails()
        
        try:
            server = poplib.POP3_SSL(HOST, PORT)
            server.user(USERNAME)
            server.pass_(PASSWORD)
            
            num_messages = len(server.list()[1])
            if num_messages == 0:
                logging.info("Nenhum e-mail na caixa de entrada.")
            
            for i in range(1, num_messages + 1):
                # Baixa o cabeçalho primeiro para obter o Message-ID sem baixar o e-mail inteiro
                # Nota: POP3 não tem um comando eficiente como o 'FETCH HEADER' do IMAP.
                # A abordagem mais simples é baixar o e-mail completo.
                response, lines, octets = server.retr(i)
                msg_content = b'\r\n'.join(lines)
                msg = BytesParser().parsebytes(msg_content)
                
                email_id = msg.get('Message-ID')
                subject = decode_header_value(msg.get('subject', 'Sem Assunto'))

                if email_id in processed_ids:
                    logging.info(f"Ignorando e-mail já processado (ID: {email_id})")
                    continue

                logging.info(f"Processando novo e-mail com assunto: '{subject}'")
                subject_lower = subject.lower()

                try:
                    task_executed = False
                    if ASSUNTO_FRETE in subject_lower:
                        logging.info("Assunto de FRETE detectado. Acionando o script de frete.")
                        processar_email_frete(msg)
                        task_executed = True

                    elif ASSUNTO_NOTA_FISCAL in subject_lower:
                        logging.info("Assunto de NOTA FISCAL detectado. Acionando o script principal.")
                        processar_email_nota_fiscal(msg)
                        task_executed = True
                    elif ASSUNTO_VEICULO in subject_lower:
                        logging.info("Assunto de DESPESA DE VEICULO detectado. Acionando o script principal.")
                        processar_email_despesa_veiculo(msg)
                        task_executed = True
                    else:
                        logging.warning(f"O assunto '{subject}' não corresponde a nenhuma ação. E-mail ignorado.")
                        # Consideramos ignorado como "processado" para não tentar de novo
                        task_executed = True

                    # Se a tarefa foi executada sem erros (ou ignorada), adiciona o ID à lista
                    if task_executed and email_id:
                        processed_ids.append(email_id)
                        save_processed_emails(processed_ids)
                        logging.info(f"E-mail (ID: {email_id}) processado com sucesso e salvo no JSON.")

                except Exception as e:
                    logging.error(f"Ocorreu um erro ao processar o e-mail com assunto '{subject}': {e}")
                    # Em caso de erro, o ID não é salvo, permitindo uma nova tentativa na próxima verificação.
            
            server.quit()

        except Exception as e:
            logging.error(f"Erro ao conectar ou comunicar com o servidor de e-mail: {e}")

        logging.info("Aguardando 5 segundos para a próxima verificação...")
        time.sleep(5)