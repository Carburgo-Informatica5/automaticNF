import os
import xml.etree.ElementTree as ET
import threading
import queue
import tkinter as tk
import os
import shutil
import time
import datetime
import pyautogui as gui
import pygetwindow as gw
import pytesseract
from email import parser
from email.header import decode_header
import os
import poplib
import json

with open('dados_extraidos_nota_39206.xml.json', 'r') as f:
    dados_json = json.load(f)

cnpj = dados_json["eminente"]["cnpj"]
nome_eminente = dados_json["eminente"]["nome"]
cnpj_destinatario = dados_json["destinatario"]["cnpj"]
nome_destinatario = dados_json["destinatario"]["nome"]
chave_acesso = dados_json["chave_acesso"]["chave"]
nmr_nota = dados_json["num_nota"]["numero_nota"]
data_emi = dados_json["data_emi"]["data_emissao"]
data_vali = dados_json["data_vali"]["data_validade"]
modelo = dados_json["modelo"]["modelo"]
valor_total = dados_json["valor_total"][0]["Valor_total"]

class SistemaNF():
    def __init__(self, master):
        self.fila_eventos = queue.Queue()
        self.br = None
        self.processamento_pausado = False

    # def iniciar_bravos(self):
    #     config = {"bravos_usr": "caetano.apollo", "bravos_pswd": "123"}
    #     self.br = openBravos.infoBravos(config, m_queue=openBravos.faker())
    #     self.br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")

    def pausar_processamento(self):
        self.processamento_pausado = True
        self.registrador.registrar("PAUSA", "Processamento de NFs pausado")
        self.interface.rotulo_status.config(text="Processamento Pausado")

    def retomar_processamento(self):
        self.processamento_pausado = False
        self.registrador.registrar_evento("CONTINUAR", "Processamento de NFs retomado")
        self.interface.rotulo_status.config(text="Processamento Retomado")

    # sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# from bravos import openBravos


# Função para ler e extrair dados do arquivo XML
def parse_nota_fiscal(xml_file_path):
    """
        Lê um arquivo XML de nota fiscal e extrai as principais informações.

    Args:
        xml_file_path (str): Caminho para o arquivo XML da nota fiscal.

    Returns:
        dict: Um dicionário com as informações extraídas da nota fiscal.
    """
    try:
        # Lê o arquivo XML
        tree = ET.parse(xml_file_path)
        root = tree.getroot()

        namespaces = {
            "ns0": "http://www.portalfiscal.inf.br/nfe",
            "ns1": "http://www.w3.org/2000/09/xmldsig#",
        }

        nota_fiscal_data = {
            "emitente": {},
            "destinatario": {},
            "produtos": [],
            "chave_acesso": {},
            "num_nota": {},
            "data_emi": {},
            "data_vali": {},
            "modelo": {},
            "valor_total": [],
            "pagamento_parcelado": [],
        }

        # Extrair informações do emitente
        emitente = root.find(".//ns0:emit", namespaces)
        if emitente is not None:
            nota_fiscal_data["emitente"] = {
                "cnpj": emitente.findtext("ns0:CNPJ", namespaces=namespaces),
                "nome": emitente.findtext("ns0:xNome", namespaces=namespaces),
            }
        else:
            print("Emitente não encontrado")

        # Extrair informações do destinatário
        destinatario = root.find(".//ns0:dest", namespaces)
        if destinatario is not None:
            nota_fiscal_data["destinatario"] = {
                "cnpj": destinatario.findtext("ns0:CNPJ", namespaces=namespaces),
                "nome": destinatario.findtext("ns0:xNome", namespaces=namespaces),
            }
        else:
            print("Destinatário não encontrado")

        # Extrair informações da nota
        num_nota = root.find(".//ns0:ide", namespaces)
        if num_nota is not None:
            nota_fiscal_data["num_nota"] = {
                "numero_nota": num_nota.findtext("ns0:nNF", namespaces=namespaces)
            }
        else:
            print("Número da nota não encontrado")

        # Extrair forma de pagamento
        forma_pagamento = root.findall(".//ns0:cobr", namespaces)
        if forma_pagamento is not None:
            for pagamento in forma_pagamento:
                parcelas = pagamento.findall("ns0:dup", namespaces=namespaces)

                if parcelas:
                    # Se existem parcelas, itere sobre elas
                    for parcela in parcelas:
                        nmr_parc = parcela.findtext("ns0:nDup", namespaces=namespaces)
                        data_venc = parcela.findtext("ns0:dVenc", namespaces=namespaces)
                        valor_parc = parcela.findtext("ns0:vDup", namespaces=namespaces)

                        if (
                            nmr_parc is not None
                            and data_venc is not None
                            and valor_parc is not None
                        ):
                            nota_fiscal_data["pagamento_parcelado"].append(
                                {
                                    "nmr_parc": nmr_parc,
                                    "data_venc": data_venc,
                                    "valor_parc": valor_parc,
                                }
                            )
            else:
                # Se não existem parcelas, verificar se há um pagamento único
                valor_total = pagamento.findtext(
                    "ns0:fat/ns0:vLiq", namespaces=namespaces
                )
                if valor_total is not None:
                    # Tratar como pagamento único
                    nota_fiscal_data["valor_total"].append(
                        {  # Pode ser considerado como a única parcela
                            "data_venc": pagamento.findtext(
                                "ns0:dup/ns0:dVenc", namespaces=namespaces
                            )
                            or "N/A",
                            "Valor_total": valor_total,
                        }
                    )
        else:
            print(
                "Forma de pagamento não encontrada"
            )  # Inicialização do dicionário nota_fiscal_data

        # Extrair data de emissão
        data_emi = num_nota.findtext("ns0:dhEmi", namespaces=namespaces)
        data_emi_format = data_emi[:10].replace("-", " ")
        if data_emi is not None:
            nota_fiscal_data["data_emi"] = {
                "data_emissao": f"{data_emi_format[8:10]}{data_emi_format[5:7]}{data_emi_format[0:4]}"
            }
        else:
            print("Data de emissão não encontrada")

        # Extrair data de validade
        data_vali = num_nota.findtext("ns0:dhSaiEnt", namespaces=namespaces)
        data_vali_format = data_vali[:10].replace("-", " ")
        if data_vali is not None:
            nota_fiscal_data["data_vali"] = {
                "data_validade": f"{data_vali_format[8:10]}{data_vali_format[5:7]}{data_vali_format[0:4]}"
            }
        else:
            print("Data de validade não encontrada")

        # Extrair modelo
        modelo = num_nota.findtext("ns0:mod", namespaces=namespaces)
        if modelo is not None:
            nota_fiscal_data["modelo"] = {"modelo": modelo}
        else:
            print("Modelo não encontrado")

        # Extrair chave de acesso
        chaveAcesso = root.find(".//ns0:infNFe", namespaces)
        if chaveAcesso is not None:
            chave_completa = chaveAcesso.get("Id")
            nota_fiscal_data["chave_acesso"] = {
                "chave": chave_completa[3:]  # Remove os 3 primeiros caracteres
            }
        else:
            print("Chave de acesso não encontrada")

        # Extrair produtos
        produtos = root.findall(".//ns0:det", namespaces)
        if produtos:
            for produto in produtos:
                prod_data = {
                    "codigo": produto.findtext(".//ns0:cProd", namespaces=namespaces),
                    "descricao": produto.findtext(
                        ".//ns0:xProd", namespaces=namespaces
                    ),
                    "quantidade": produto.findtext(
                        ".//ns0:qCom", namespaces=namespaces
                    ),
                    "valor_total_prod": produto.findtext(
                        ".//ns0:vProd", namespaces=namespaces
                    ),
                }
                nota_fiscal_data["produtos"].append(prod_data)
        else:
            print("Produtos não encontrados")

        return nota_fiscal_data
    except ET.ParseError as e:
        print(f"Erro ao parsear o arquivo XML: {e}")
        return None


def processar_notas_fiscais(xml_folder):
    """
        Processa arquivos XML de notas fiscais, insere os dados no banco,
        realiza o lançamento no sistema Bravos e envia um relatório por email.

    Args:
        xml_folder (str): Diretório onde os arquivos XML das notas fiscais estão localizados.
    """

    resultados = []

    # Função para abrir o sistema

    for xml_file in os.listdir(xml_folder):
        if xml_file.endswith(".xml"):
            xml_path = os.path.join(xml_folder, xml_file)

            # Extrair dados do XMl
            dados_nf = parse_nota_fiscal(xml_path)
            if dados_nf:
                # Lançar nota no sistema Bravos
                sucesso_bravos = inserir_dados_no_bravos(dados_nf)

                if sucesso_bravos:
                    mover_nota(xml_path, xml_folder)

                resultados.append(
                    {
                        "arquivo": xml_file,
                        "status": "Sucesso" if sucesso_bravos else "Falha",
                        "detalhes": (
                            "Lançamento bem-sucedido"
                            if sucesso_bravos
                            else "Erro ao lançar nota"
                        ),
                    }
                )
            else:
                resultados.append(
                    {
                        "arquivo": xml_file,
                        "status": "Falha",
                        "detalhes": "Erro ao ler o XML",
                    }
                )

    # Função fechar o sistema Bravos

    enviar_relatorio_email(resultados)


def mover_nota(xml_path, xml_folder):
    """
    Move a nota fiscal processada para a pasta 'processadas'

    Args:
        xml_path (str): Caminho completo do arquivo XML processado
        xml_folder (str): Diretório base dos XMLs
    """

    # Cria a pasta 'processadas' se não existir
    pasta_processadas = os.path.join(xml_folder, "processadas")
    os.makedirs(pasta_processadas, exist_ok=True)

    # Nome do arquivo Original
    nome_arquivo = os.path.basename(xml_path)

    # Caminho do destino
    destino = os.path.join(pasta_processadas, nome_arquivo)

    try:
        # Move o arquivo
        shutil.move(xml_path, destino)
        print(f"Arquivo {nome_arquivo} movido para pasta processadas")
    except Exception as e:
        print(f"Erro ao mover a nota do fiscal {nome_arquivo}: {e}")


# Função auxiliar para inserir dados no sistema Bravos
def inserir_dados_no_bravos(dados_nf):
    """
    Insere os dados da nota fiscal no sistema Bravos.

    Args:
        dados_nf (dict): Dados da nota fiscal extraídos do XML.

    Returns:
        bool: Retorna True se a inserção foi bem-sucedida, caso contrário, False.
    """

    # Implementação do fluxo para inserir no Sistema Bravos via pyautogui
    # Configurações do servidor e credenciais
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
    
    def save_attachment(part, directory):
        filename = decode_header_value(part.get_filename())
        if not filename:
            filename = "untitled"
        if not filename.lower().endswith(".xml"):
            filename += ".xml"
        if not os.path.exists(directory):
            os.makedirs(directory)
        filepath = os.path.join(directory, filename)
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))
        print(f"Anexo salvo em: {filepath}")
        
    def processar_centros_de_custo(cc_texto):
        """
        Processa o texto extraído para identificar se é valor ou porcentagem e retorna
        o formato adequado para o preenchimento.
        """
        centros_de_custo = []

        # Divida os centros de custo por vírgulas
        for item in cc_texto.split(","):
            item = item.strip()  # Remove espaços extras no início e no final

            # Verifique se contém '%' indicando porcentagem
            if "%" in item:
                # A separação entre centro de custo e porcentagem deve ser feita com '-'
                try:
                    # Remove os espaços ao redor do hífen antes de separar
                    cc, porcentagem = [x.strip() for x in item.split("-")]
                    # Verifica se a porcentagem está no formato correto
                    if porcentagem.replace("%", "").strip().isnumeric():
                        centros_de_custo.append(
                            (cc, None, porcentagem)
                        )  # Adiciona como porcentagem
                    else:
                        print(f"Formato de porcentagem inválido: {porcentagem}")
                except ValueError:
                    print(f"Erro ao processar centro de custo com porcentagem: {item}")
            else:
                try:
                    # Remove os espaços ao redor do hífen antes de separar
                    cc, valor = [x.strip() for x in item.split("-")]
                    # Verifica se o valor está no formato correto (numérico)
                    if valor.replace(".", "", 1).isdigit():
                        centros_de_custo.append(
                            (cc, valor, None)
                        )  # Adiciona como valor monetário
                    else:
                        print(f"Formato de valor inválido: {valor}")
                except ValueError:
                    print(f"Erro ao processar centro de custo com valor: {item}")

        return centros_de_custo
    
    dados_centros_de_custo = []
    
    try:
        server = poplib.POP3_SSL(HOST, PORT)
        server.user(USERNAME)
        server.pass_(PASSWORD)
        num_messages = len(server.list()[1])
        print(f"Você tem {num_messages} mensagem(s) no servidor.")

        for i in range(num_messages):
            response, lines, octets = server.retr(i + 1)
            raw_message = b"\n".join(lines).decode("utf-8", errors="ignore")
            email_message = parser.Parser().parsestr(raw_message)
            subject = decode_header_value(email_message["subject"])

            if subject == ASSUNTO_ALVO:
                print(f"\nE-mail encontrado com o assunto: {subject}")
                body = ""
                if email_message.is_multipart():
                    for part in email_message.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/plain":
                            charset = part.get_content_charset()
                            body = decode_body(part.get_payload(decode=True), charset)
                        elif part.get("Content-Disposition") is not None:
                            save_attachment(part, DIRECTORY)
                else:
                    charset = email_message.get_content_charset()
                    body = decode_body(email_message.get_payload(decode=True), charset)
                valores_extraidos = extract_values(body)
                departamento = valores_extraidos["departamento"]
                origem = valores_extraidos["origem"]
                descricao = valores_extraidos["descricao"]
                revenda_cc = valores_extraidos["revenda_cc"]
                cc_texto = valores_extraidos["cc"]
                rateio = valores_extraidos["rateio"]
                cod_item = valores_extraidos["cod_item"]
                dados_centros_de_custo = processar_centros_de_custo(cc_texto)
                break
        server.quit()

    except Exception as e:
        print(f"Erro: {e}")

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
    x, y = window.left + 275, window.top + 80  # Ajuste os offsets conforme necessário
    time.sleep(3)
    gui.moveTo(x, y, duration=0.5)
    gui.click()
    time.sleep(5)
    gui.press("tab", presses=19)
    gui.write(cnpj)
    time.sleep(2)
    gui.press("enter")

    pytesseract.pytesseract_cmd = r"C:\Program Files\Tesseract-OCR/tesseract.exe"

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
    gui.write(cod_item)  # variavel via email "Outros"
    gui.press("tab", presses=10)
    gui.write("1")
    gui.press("tab")
    gui.write(valor_total)
    gui.press("tab", presses=26)
    gui.write(descricao)  # Variavel Relativa puxar pelo email
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
    for cc, valor, porcentagem in dados_centros_de_custo:
        gui.write(revenda_cc)
        gui.press("tab", presses=3)
        gui.write(cc)
        gui.press("tab", presses=2)
        gui.write(origem)
        gui.press("tab", presses=2)
        if valor:
            gui.write(valor)
            gui.press("tab", presses=3)
            gui.press(["f2", "f3"])
        elif porcentagem:
            gui.press("tab")
            gui.write(porcentagem)
            gui.press("tab", presses=3)
            gui.press(["f2", "f3"])
        gui.press("tab", presses=19)
    gui.moveTo(1271, 758)
    gui.click()
    gui.press("tab", presses=3)
    gui.press("enter")


# Função para enviar relatório para o email
def enviar_relatorio_email(resultados):
    """
    Envia um relatório de status de cada nota fiscal processada por email.

    Args:
        resultados (list): Lista de dicionários com o status de cada nota fiscal.
    """
    mensagem = "Relatório de Lançamento de Notas Fiscais: \n\n"
    for resultado in resultados:
        mensagem += f"Arquivo: {resultado['arquivo']}, Status: {resultado["status"]}, Detalhes: {resultado["detalhes"]}\n"

        # Função para enviar o relatório por email

    def executar(self):
        threading.Thread(target=self.interface.run, daemon=True).start()
        self.iniciar_bravos()

        # Inicia o processamento em uma thread separada
        threading.Thread(
            target=self.processar_notas_fiscais, args=("xml_folder",), daemon=True
        ).start()

        # Loop principal para atualizar a interface
        while True:
            try:
                evento = self.fila_eventos.get(timeout=0.1)
                self.interface.texto_log.insert(
                    tk.END, f"{evento['timestamp']} - {evento['descricao']}\n"
                )
                self.interface.texto_log.see(tk.END)
            except queue.Empty:
                pass

            self.interface.update()  # Atualiza a interface


if __name__ == "__main__":
    root = tk.Tk()
    sistema = SistemaNF(root)
    sistema.executar()
    root.mainloop()