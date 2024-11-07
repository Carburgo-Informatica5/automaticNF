import os
import xml.etree.ElementTree as ET
import threading
import queue
import tkinter as tk
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bravos import openBravos

config = {"bravos_usr": "seu_usuario", "bravos_pswd": "sua_senha"}
br = openBravos.infoBravos(config, m_queue=openBravos.faker())

br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")

# Função para ler e extrair dados do arquivo XML
def  parse_nota_fiscal  (xml_file_path):
    """
        Lê um arquivo XML de nota fiscal e extrai as principais informações.
    
    Args:
        xml_file_path (str): Caminho para o arquivo XML da nota fiscal.
        
    Returns:
        dict: Um dicionário com as informações extraídas da nota fiscal.
    """
    
    try: 
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        nota_fiscal_data = {
            "emitente": {},
            "destinatario": {},
            "produtos": [],
            "chave_acesso": {},
            "num_nota": {},
            "data_emi": {},''
            "data_vali": {},
            "modelo": {},
        }
        
        # Extrair informações eminente e destinatario
        eminente = root.find('.//emit')
        if eminente is not None:
            nota_fiscal_data["eminente"] = {
                "cnpj": eminente.findtext("CNPJ"),
                "nome": eminente.findtext("xNome")
            }
            
        destinatario = root.find('.//dest')
        if destinatario is not None:
            nota_fiscal_data["destinatario"] = {
                "cnpj":  destinatario.findtext("CNPJ"),
                "nome": destinatario.findtext("xNome")
            }
            
            # Extrair infos para cadastro de notas
            num_nota = root.find('.//ide')
            if num_nota is not None:
                nota_fiscal_data["num_nota"] = {
                    "numero_nota": num_nota.findtext("nNF")
                }

            data_emi = root.find('.//ide')
            if data_emi is not None:
                nota_fiscal_data["data_emi"] = {
                    "data_emissao": data_emi.findtext("dhEmi")
                    # Tem que ver como pegar e converter a data
                }

            data_vali = root.find('.//ide')
            if data_vali is not None:
                nota_fiscal_data["data_vali"] = {
                    "data_validade": data_vali.findtext("dhSaiEnt")
                    # Tem que ver como pegar e converter a data
                }
                
            modelo = root.find(".//ide")
            if modelo is not None:
                nota_fiscal_data["modelo"] = {
                    "modelo": modelo.findtext("mod")
                }

            chaveAcesso = root.find('.//infNFe') 
            if chaveAcesso is not None:
                nota_fiscal_data["chaveAcesso"] = {
                    "chave": chaveAcesso.findtext("Id")
                    # Pegar apenas a numeração
                }
            
        # Extrair Produto
        produtos = root.find('.//det')
        for produto in produtos:
            prod_data = {
                "codigo": produto.findtext(".//cProd"),
                "descricao": produto.findtext(".//xProd"),
                "quantidade":  produto.findtext(".//qProd"),
                "valor_unitario": produto.findtext(".//vUnid"),
                "valor_total": produto.findtext(".//vProd")
            }
            
            nota_fiscal_data["produtos"].append(prod_data)
        
        return nota_fiscal_data
    
    except ET.ParseError as e:
        print("Erro ao processar o XML:", e)
        return None
    
def processar_notas_fiscais(xml_folder):
    """
        Processa arquivos XML de notas fiscais, insere os dados no banco,
        realiza o lançamento no sistema Bravos e envia um relatório por email.

    Args:
        xml_folder (str): Diretório onde os arquivos XML das notas fiscais estão localizados.
    """
    
    resultados  = []
    
    # Função para abrir o sistema
        
    for xml_file in os.listdir(xml_folder):
        if xml_file.endswith('.xml'):
            xml_path = os.path.join(xml_folder, xml_file)
            
            # Extrair dados do XMl
            dados_nf = parse_nota_fiscal(xml_path)
            if dados_nf:
                #Lançar nota no sistema Bravos
                sucesso_bravos = inserir_dados_no_bravos(dados_nf)
                
                resultados.append({
                    "arquivo": xml_file,
                    "status": "Sucesso" if sucesso_bravos else "Falha",
                    "detalhes": "Lançamento bem-sucedido" if sucesso_bravos else "Erro ao lançar nota"
                })
            else:
                resultados.append({
                    "arquivo": xml_file,
                    "status": "Falha",
                    "detalhes": "Erro ao ler o XML"
                })
                
    # Função fechar o sistema Bravos
    
    enviar_relatorio_email(resultados)
    
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
    try:
        # Exemplo básico (Substituir pelo caminho real para cada campo)
        # Inserir cada campo no Bravos
        return True
    except Exception as e:
        print("Erro ao inserir dados no Bravos:", e)
        return False

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
        
from interface import interfaceMonitoramentoNF
from tratamentoErros import tratadorErros
from sistemaLogs import RegistradorSistema

class SistemaNF:
    def __init__(self):
        self.fila_eventos = queue.Queue()
        self.interface = interfaceMonitoramentoNF()
        self.tratador_erros = tratadorErros()
        self.registrador = RegistradorSistema()
        self.br = None

    def iniciar_bravos(self):
        config = {"bravos_usr": "seu_usuario", "bravos_pswd": "sua_senha"}
        self.br = openBravos.infoBravos(config, m_queue=openBravos.faker())
        self.br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")

    def processar_notas_fiscais(self, xml_folder):
        try:
            self.registrador.registrar_evento("INICIO", "Iniciando processamento de notas fiscais")
            # Aqui você coloca a lógica de processamento das notas fiscais
            # Use self.fila_eventos.put() para enviar atualizações para a interface
            # Exemplo:
            self.fila_eventos.put({"timestamp": "2023-05-20 10:00:00", "descricao": "Processando nota fiscal X"})
            
            # Simula o processamento
            # Substitua isso pela sua lógica real de processamento
            import time
            time.sleep(5)
            
            self.registrador.registrar_evento("FIM", "Processamento de notas fiscais concluído")
        except Exception as e:
            self.tratador_erros.tratarErros(e, "Processamento de notas fiscais")

    def executar(self):
        threading.Thread(target=self.interface.run, daemon=True).start()
        self.iniciar_bravos()
        
        # Inicia o processamento em uma thread separada
        threading.Thread(target=self.processar_notas_fiscais, args=("caminho/para/pasta/xml",), daemon=True).start()

        # Loop principal para atualizar a interface
        while True:
            try:
                evento = self.fila_eventos.get(timeout=0.1)
                self.interface.texto_log.insert(tk.END, f"{evento['timestamp']} - {evento['descricao']}\n")
                self.interface.texto_log.see(tk.END)
            except queue.Empty:
                pass
            
            self.interface.update()  # Atualiza a interface

if __name__ == "__main__":
    sistema = SistemaNF()
    sistema.executar()