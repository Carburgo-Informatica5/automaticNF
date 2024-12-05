import poplib
from email import parser
from email.header import decode_header
import os

# Configurações do servidor e credenciais
HOST = 'mail.carburgo.com.br'
PORT = 995
USERNAME = 'caetano.apollo@carburgo.com.br'
PASSWORD = 'p@r!sA1856'

ASSUNTO_ALVO = 'Lançamentos notas fiscais DANI'

DIRECTORY = 'anexos'

def decode_header_value(header_value):
    decoded, encoding = decode_header(header_value)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else 'utf-8', errors='ignore')
    return decoded

def decode_body(payload, charset):
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
            values['cod_item'] = line.split(":", 1)[1].strip()

    return values

def process_rateio(rateio):
    """Processa a string de rateio e retorna uma lista de pares (cc, valor)."""
    rateio_items = rateio.split(',')
    parsed_rateio = []

    for item in rateio_items:
        try:
            cc, valor = item.split('-')
            cc = cc.strip()  # Centro de custo
            valor = valor.strip()  # Valor ou porcentagem
            parsed_rateio.append((cc, valor))
        except ValueError:
            print(f"Erro ao processar item de rateio: {item}")
            continue

    return parsed_rateio

try:
    # Conectar ao servidor POP3 com SSL
    server = poplib.POP3_SSL(HOST, PORT)
    server.user(USERNAME)
    server.pass_(PASSWORD)

    num_messages = len(server.list()[1])
    print(f'Você tem {num_messages} mensagem(s) no servidor.')

    for i in range(num_messages):
        response, lines, octets = server.retr(i + 1)
        raw_message = b'\n'.join(lines).decode('utf-8', errors='ignore')
        email_message = parser.Parser().parsestr(raw_message)

        subject = decode_header_value(email_message['subject'])

        if subject == ASSUNTO_ALVO:
            print(f"\nE-mail encontrado com o assunto: {subject}")
            body = ""
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        charset = part.get_content_charset()
                        body = decode_body(part.get_payload(decode=True), charset)
            else:
                charset = email_message.get_content_charset()
                body = decode_body(email_message.get_payload(decode=True), charset)

            valores_extraidos = extract_values(body)
            departamento = valores_extraidos['departamento']
            origem = valores_extraidos['origem']
            descricao = valores_extraidos['descricao']
            revenda_cc = valores_extraidos['revenda_cc']
            cc = valores_extraidos['cc']
            rateio = valores_extraidos['rateio']
            cod_item = valores_extraidos['cod_item']
            
            print(f"Departamento: {departamento}")
            print(f"Origem: {origem}")
            print(f"Descrição: {descricao}")
            print(f"CC: {cc}")
            print(f"Rateio: {rateio}")
            print(f"código de tributação: {cod_item}")
            break

    server.quit()

except Exception as e:
    print('Erro:', e)