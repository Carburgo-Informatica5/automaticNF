import poplib
from email import parser
from email.header import decode_header
import os

# Configurações do servidor e credenciais
HOST = 'mail.carburgo.com.br'  # Servidor POP3 corporativo
PORT = 995                    # Porta segura (SSL)
USERNAME = 'caetano.apollo@carburgo.com.br'
PASSWORD = 'p@r!sA1856'

# Assunto alvo para busca
ASSUNTO_ALVO = 'Lançamentos notas fiscais DANI'

# Diretório para salvar os anexos
DIRECTORY = 'anexos'  # Pasta onde os anexos serão salvos

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
        'percentual': None
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
            cc_values = line.split(":", 1)[1].strip()
            values['cc'] = [x.strip() for x in cc_values.replace(',', ' ').split()]
        elif line.startswith("rateio:"):
            values['rateio'] = line.split(":", 1)[1].strip()
        elif line.startswith("percentual:"):
            percentual_values = line.split(":", 1)[1].strip()
            values['percentual'] = [x.strip() for x in percentual_values.replace(',', ' ').split()]

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
            cc = valores_extraidos['cc']
            rateio = valores_extraidos['rateio']
            percentual = valores_extraidos['percentual']
            
            print(f"Departamento: {departamento}")
            print(f"Origem: {origem}")
            print(f"Descrição: {descricao}")
            print(f"CC: {cc}")
            print(f"Rateio: {rateio}")
            print(f"Percentual: {percentual}")

            break  # Para após encontrar o primeiro e-mail com o assunto desejado

    # Desconectar do servidor
    server.quit()

except Exception as e:
    print('Erro:', e)