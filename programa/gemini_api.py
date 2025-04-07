import time
import google.generativeai as genai
import unicodedata
import requests
import logging
import json

class GeminiAPI:
    def __init__(self, api_key):
        # Configura a chave da API do Gemini
        genai.configure(api_key=api_key)

    def normalize_text(self, text):
        """Normaliza o texto para remover caracteres especiais e problemas de codificação."""
        return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')

    def get_cnpj_by_razao_social(self, razao_social):
        """Consulta o CNPJ pela razão social utilizando a API ReceitaWS"""
        url = f"https://www.receitaws.com.br/v1/cnpj/{razao_social}"
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            return data.get('cnpj', None)
        else:
            print(f"Erro ao consultar CNPJ para '{razao_social}': {response.status_code}")
            return None

    def upload_pdf(self, path, mime_type="application/pdf"):
        """Faz o upload do arquivo PDF para o Gemini."""
        file = genai.upload_file(path, mime_type=mime_type)
        print(f"Arquivo '{file.display_name}' carregado com sucesso como: {file.uri}")
        return {"success": True, "file_id": file.name}

    def check_processing_status(self, file_id):
        """Verifica o status de processamento do arquivo até ele ser ativo."""
        print("Aguardando o processamento do arquivo...")
        file = genai.get_file(file_id)
        while file.state.name == "PROCESSING":
            print(".", end="", flush=True)
            time.sleep(10)
            file = genai.get_file(file_id)
        if file.state.name != "ACTIVE":
            raise Exception(f"Arquivo {file.name} falhou ao processar.")
        print("Arquivo pronto para uso.")
        return {"state": file.state.name}

    def identify_document_type(self, text):
        """Tenta identificar o tipo de documento com base em palavras-chave."""
        if "nota fiscal" in text.lower():
            return "Nota Fiscal"
        elif "fatura" in text.lower():
            return "Fatura"
        elif "boleto" in text.lower():
            return "Boleto"
        else:
            return "Desconhecido"

    def extract_info(self, file_id):
        generation_config = {
            "temperature": 0.1,
            "top_p": 0.95,
            "top_k": 40,
            "response_mime_type": "text/plain",
        }

        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=generation_config,
        )

        file_obj = genai.get_file(file_id)
        chat_session = model.start_chat(history=[])

        response = chat_session.send_message([
        {"text": """
        Você receberá um arquivo PDF contendo informações de um documento financeiro. Sua tarefa é extrair os dados relevantes e retornar um JSON bem formatado e válido. Certifique-se de seguir as instruções abaixo:

        1. Extraia as seguintes informações, se disponíveis:
            - Número da nota fiscal ou número da conta (apenas números, sem caracteres especiais ou traços).
            - Data de emissão (formato DDMMYYYY, sem barras ou outros caracteres).
            - Nome completo do prestador de serviço.
            - Nome completo do tomador de serviço.
            - CNPJ do prestador de serviço (apenas números, sem pontos ou traços). Caso o CNPJ não esteja presente, retorne "Não encontrado".
            - CNPJ do tomador de serviço (apenas números, sem pontos ou traços). Caso o CNPJ não esteja presente, retorne "Não encontrado".
            - Endereço do tomador de serviço (Somente número e cidade).
            - Valor total da nota fiscal (substitua vírgulas por pontos para valores decimais).
            - Valor líquido (substitua vírgulas por pontos para valores decimais).
            - Verifique se há ISS retido: "Sim" ou "Não". Caso esteja presente, inclua o campo "ISS Retido": "Sim". Caso contrário, "ISS Retido": "Não".
            - Valores dos seguintes impostos, se presentes. Caso algum imposto não esteja presente, retorne "0.00" para esse imposto:
                - PIS
                - COFINS
                - INSS
                - ISS Retido (valor real, não o ISS total)
                - IR
                - CSLL
            - Tipo de documento (identifique se é uma "Nota Fiscal", "Boleto" ou "Fatura").

        2. Caso o PDF contenha informações ilegíveis ou ausentes, retorne um campo indicando "Dados não legíveis" ou "Informação ausente".

        3. Certifique-se de que o JSON esteja bem formatado, sem caracteres inválidos ou erros de sintaxe.

        Exemplo de JSON esperado:

        {
            "emitente": {
                "cnpj": "12345678000195",
                "nome": "Empresa Prestadora de Serviços LTDA"
            },
            "destinatario": {
                "cnpj": "98765432000112",
                "nome": "Empresa Tomadora de Serviços SA"
                "endereco": "123, Cidade Exemplo"
            },
            "num_nota": "12345",
            "data_emissao": "01042025",
            "valor_total": "1500.00",
            "valor_liquido": "1400.00",
            "ISS_retido": "Sim",
            "impostos": {
                "PIS": "50.00",
                "COFINS": "100.00",
                "INSS": "0.00",
                "ISS_retido": "50.00",
                "IR": "0.00",
                "CSLL": "0.00"
            },
            "tipo_documento": "Nota Fiscal"
        }

        Retorne apenas o JSON como resposta, sem explicações adicionais.
            """},
            file_obj
    ])

        # Processa a resposta do modelo
        extracted_text = self.normalize_text(response.text)
        logging.info(f"Texto extraído do PDF: {extracted_text}")

        try:
            # Tenta converter o texto extraído para JSON
            extracted_json = json.loads(extracted_text)
            
            if not isinstance(extracted_json, dict):
                logging.error("Erro: O JSON extraído não é um dicionário.")
                return None
            
            tomador = extracted_json.get("destinatario", {})
            cnpj_tomador = tomador.get("cnpj", "Não encontrado")
            if cnpj_tomador == "Não encontrado":
                razao_social = tomador.get("nome", "")
                endereco = tomador.get("endereco", "")
                cidade = tomador.get("cidade", "")

                # Concatena as informações para buscar o CNPJ
                busca_cnpj = f"{razao_social}, {endereco}, {cidade}"
                logging.info(f"Buscando CNPJ do tomador com as informações: {busca_cnpj}")
                cnpj_tomador = self.get_cnpj_by_razao_social(busca_cnpj)

            # Atualiza o JSON com o CNPJ encontrado
            if cnpj_tomador:
                tomador["cnpj"] = cnpj_tomador
            else:
                logging.warning("Não foi possível encontrar o CNPJ do tomador.")

            logging.info(f"JSON extraído e atualizado: {extracted_json}")
            return extracted_json
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao decodificar JSON: {e}")
            raise ValueError("O texto extraído não é um JSON válido.")