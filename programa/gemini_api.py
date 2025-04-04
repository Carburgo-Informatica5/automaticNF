import os
import time
import google.generativeai as genai
import unicodedata
import requests

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
        """Extrai as informações do PDF usando Gemini."""
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

        # Envia uma mensagem detalhada de extração
        response = chat_session.send_message([
            {"text": """
Por favor, extraia as seguintes informações de uma nota fiscal em PDF e retorne os dados em formato JSON. 
Certifique-se de processar corretamente os dados, especialmente considerando que a nota pode conter diferentes tipos de informação, como valores de impostos e informações fiscais. Caso algum dado esteja ausente, substitua com o valor '0.00' ou 'Não', conforme indicado. Detalhe as etapas de extração de forma clara:

1. Número da nota fiscal ou Número da Conta (Apenas números, sem caracteres especiais e traços)
2. Data da emissão (formato DDMMYYYY, sem barras ou outros caracteres)
3. Nome completo do prestador de serviço
4. Nome completo do tomador do serviço
5. CNPJ do prestador de serviço (apenas números, sem pontos ou traços) se não houver CNPJ quero que você pesquise a razão social e retorne o CNPJ
6. CNPJ do tomador de serviço (apenas números, sem pontos ou traços) se não houver CNPJ quero que você pesquise a razão social e retorne o CNPJ
7. Valor total da nota fiscal (substitua vírgulas por pontos para valores decimais)
8. Valor líquido (substitua vírgulas por pontos para valores decimais)
9. Verifique se há ISS retido: 'Sim' ou 'Não'. Caso esteja presente, inclua o campo 'ISS Retido': 'Sim'. Caso contrário, 'ISS Retido': 'Não'.
10. Extração dos valores dos seguintes impostos, se presentes. Caso algum imposto não esteja presente, retorne '0.00' para esse imposto:
    - PIS
    - COFINS
    - INSS
    - ISS Retido (valor real, não o ISS total)
    - IR
    - CSLL
11. Certifique-se de que o JSON esteja corretamente formatado, sem caracteres inválidos ou faltantes. A formatação deve ser limpa e seguir a estrutura solicitada.
12. Caso o PDF contenha imagens ou informações ilegíveis, retorne um campo indicando 'Dados não legíveis' ou 'Informação ausente' conforme o caso.
13. Adicione também o campo 'Tipo de Documento' para identificar se é uma nota fiscal, boleto ou fatura, caso haja esse tipo de informação disponível no PDF.

Exemplo de JSON esperado:

{
    "emitente": {
        "cnpj": "27249934000118",
        "nome": "EDITORAZANE EDITORA EIRELLI ME"
    },
    "destinatario": {
        "cnpj": "04682292000655",
        "nome": "EIFFEL VEICULOS COMERCIO E IMPORTAÇÃO LTDA"
    },
    "num_nota": {
        "numero_nota": "112"
    },
    "data_venc": {
        "data_venc": "31032025"
    },
    "data_emi": {
        "data_emissao": "12122024"
    },
    "valor_total": {
        "valor_total": "1000.00"
    },
    "valor_liquido": {
        "valor_liquido": "1000.00"
    },
    "modelo": {
        "modelo": "01"
    },
    "serie": "1",
    "chave_acesso": {
        "chave": ""
    },
    "impostos": {
        "ISS_retido": "0.00",
        "PIS": "0.00",
        "COFINS": "0.00",
        "INSS": "0.00",
        "IR": "0.00",
        "CSLL": "0.00"
    }
}
            """},
            file_obj
        ])

        # Processa a resposta do modelo
        extracted_data = self.normalize_text(response.text)
        
        # Identifica o tipo de documento
        document_type = self.identify_document_type(extracted_data)
        
        # Verifica se a extração do CNPJ do tomador está ausente e tenta buscar pelo nome
        cnpj_tomador = None
        if "destinatario" in extracted_data and "nome" in extracted_data["destinatario"]:
            razao_social_tomador = extracted_data["destinatario"]["nome"]
            cnpj_tomador = self.get_cnpj_by_razao_social(razao_social_tomador)
        
        if cnpj_tomador:
            extracted_data["destinatario"]["cnpj"] = cnpj_tomador
        else:
            extracted_data["destinatario"]["cnpj"] = "Não encontrado"
        
        # Retorna o JSON de resposta
        return extracted_data