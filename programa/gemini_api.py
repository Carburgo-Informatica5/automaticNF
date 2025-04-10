import time
import google.generativeai as genai
import unicodedata
import logging
import json
import re

from db_connection import cnpj

class GeminiAPI:
    def __init__(self, api_key):
        genai.configure(api_key=api_key)

    def normalize_text(self, text):
        """Normaliza o texto para remover caracteres especiais."""
        return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')

    def upload_pdf(self, path, mime_type="application/pdf"):
        """Faz o upload do arquivo PDF para o Gemini."""
        file = genai.upload_file(path, mime_type=mime_type)
        print(f"Arquivo '{file.display_name}' carregado com sucesso como: {file.uri}")
        return {"success": True, "file_id": file.name}

    def check_processing_status(self, file_id):
        """Verifica o status de processamento do arquivo."""
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
                - Número do Endereço do tomador de serviço (Somente número).
                - Cidade do Endereço do tomador de serviço (Somente cidade).
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
                    "nome": "Empresa Tomadora de Serviços SA",
                    "Numero": "123",
                    "Cidade": "Novo Hamburgo"
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
        
        logging.info(f"Resposta da API recebida. {response.text}")

        # Garantindo que temos uma string para trabalhar
        if not isinstance(response.text, str):
            logging.error("Resposta da API não é uma string")
            return None

        # Limpeza inicial do texto
        extracted_text = self.normalize_text(response.text)
        extracted_text = extracted_text.strip()
        
        # Removendo possíveis marcadores de código (```json)
        extracted_text = re.sub(r'^```json|```$', '', extracted_text, flags=re.IGNORECASE).strip()

        # Conversão para JSON/dicionário
        try:
            extracted_json = json.loads(extracted_text)
            if not isinstance(extracted_json, dict):
                logging.error("JSON extraído não é um dicionário")
                return None
        except json.JSONDecodeError as e:
            logging.error(f"Falha ao decodificar JSON: {e}")
            logging.error(f"Conteúdo problemático: {extracted_text}...")
            return None

        # Busca de CNPJ se necessário
        tomador = extracted_json.get("destinatario", {})
        cnpj_tomador = tomador.get("cnpj", "Nao encontrado")

        if cnpj_tomador == "Nao encontrado":
            try:
                num_endereco = tomador.get("Numero", "")
                cidade = tomador.get("Cidade", "")
                if num_endereco and cidade:
                    result = cnpj(num_endereco=num_endereco, cidade=cidade)
                    cnpj_tomador = result if result else "Não encontrado"
            except Exception as e:
                logging.error(f"Erro ao buscar CNPJ: {e}")
                cnpj_tomador = "Erro na consulta"

        extracted_json["destinatario"]["cnpj"] = cnpj_tomador
        return extracted_json