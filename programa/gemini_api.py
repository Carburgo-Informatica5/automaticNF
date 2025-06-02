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
            "temperature": 0,
            "top_p": 0.95,
            "top_k": 40,
            "response_mime_type": "text/plain",
        }

        model = genai.GenerativeModel(
            model_name="gemini-2.0-flash",
            generation_config=generation_config,
        )

        file_obj = genai.get_file(file_id)
        chat_session = model.start_chat(history=[])

        response = chat_session.send_message([
            {"text": """
            Você receberá um arquivo PDF de um documento fiscal (nota fiscal, fatura ou boleto). Extraia as informações abaixo e retorne **apenas** um JSON válido, sem explicações, comentários ou texto adicional.

**Campos obrigatórios do JSON:**

- "emitente":  
    - "cnpj": CNPJ do prestador (apenas números, sem pontos ou traços; se não houver, use "Nao encontrado")
    - "nome": Nome completo do prestador

- "destinatario":  
    - "cnpj": CNPJ do tomador (apenas números, sem pontos ou traços; se não houver, use "Nao encontrado")
    - "nome": Nome completo do tomador
    - "Numero": Número do endereço do tomador (apenas números; se não houver, use "Informacao ausente")
    - "Cidade": Cidade do endereço do tomador (apenas nome da cidade; se não houver, use "Informacao ausente")

- "num_nota": Número da nota fiscal, conta ou fatura (apenas números, sem caracteres especiais; priorize o número da fatura se houver)
- "data_emissao": Data de emissão (formato DDMMYYYY, sem barras ou outros caracteres; se não houver, use "Informacao ausente")
- "valor_total": Valor total da nota (use ponto como separador decimal; se não houver, use "0.00")
- "valor_liquido": Valor líquido (use ponto como separador decimal; se não houver, use o valor total)
- "ISS_retido": "Sim" se houver ISS retido, "Nao" caso contrário
- "serie": Série da nota (se não houver, use "Informacao ausente")
- "chave_acesso": Chave de acesso (apenas para notas de frete; se não houver, use "Informacao ausente")

- "impostos":  
    - "PIS": Valor do PIS (padrão "0.00" se não houver)
    - "COFINS": Valor do COFINS (padrão "0.00" se não houver)
    - "INSS": Valor do INSS (padrão "0.00" se não houver)
    - "ISS_retido": Valor do ISS retido (padrão "0.00" se não houver)
    - "IR": Valor do IR (padrão "0.00" se não houver)
    - "CSLL": Valor do CSLL (padrão "0.00" se não houver)

**Regras adicionais:**
- Se o documento for de frete, use o remetente como tomador e preencha "chave_acesso" e "série" que sempre estará ao lado esquerdo do número da nota.
- Se algum campo não for encontrado ou estiver ilegível, use "Informacao ausente" ou "Nao encontrado" conforme o caso.
- O JSON deve ser bem formatado, válido e não conter caracteres inválidos.
- Não inclua explicações, apenas o JSON.

**Exemplo de resposta esperada:**

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
    "serie": "7",
    "chave_acesso": "12345678901234567890123456789012345678901234",
    "impostos": {
        "PIS": "50.00",
        "COFINS": "100.00",
        "INSS": "0.00",
        "ISS_retido": "50.00",
        "IR": "0.00",
        "CSLL": "0.00"
    }
}

Retorne **apenas** o JSON.
            """},
            file_obj
        ])
        
        logging.info(f"Resposta da API recebida. {response.text}")

        # Garantindo que temos uma string para trabalhar
        if not isinstance(response.text, str):
            logging.error("Resposta da API não é uma string")
            return None

        try:
            # Limpeza inicial do texto
            extracted_text = self.normalize_text(response.text)
            extracted_text = extracted_text.replace("\n", " ").replace("\r", "")

            # Removendo possíveis marcadores de código (```json)
            extracted_text = re.sub(r'^```json|```$', '', extracted_text, flags=re.IGNORECASE)

            # Verificar se há múltiplos objetos JSON concatenados
            if extracted_text.count("{") > 1 and extracted_text.count("}") > 1:
                logging.warning("Texto contém múltiplos objetos JSON. Tentando corrigir...")
                extracted_text = extracted_text[extracted_text.find("{"):extracted_text.rfind("}") + 1]

            try:
                extracted_json = json.loads(extracted_text)
                if not isinstance(extracted_json, dict):
                    logging.error("JSON extraído não é um dicionário")
                    return None

                # Log para verificar o JSON extraído
                logging.info(f"JSON extraído: {extracted_json}")

                # Garantir que valor_total esteja presente
                if "valor_total" not in extracted_json or not extracted_json["valor_total"]:
                    logging.warning("Campo 'valor_total' ausente ou vazio no JSON extraído.")
                    extracted_json["valor_total"] = "0.00"  # Valor padrão
            except json.JSONDecodeError as e:
                logging.error(f"Falha ao decodificar JSON: {e}")
                logging.error(f"Conteúdo problemático: {extracted_text}")
                return None
        except json.JSONDecodeError as e:
            logging.error(f"Falha ao decodificar JSON: {e}")
            logging.error(f"Conteúdo problemático: {extracted_text}")
            return None

        # Atualizar o CNPJ se necessário
        tomador = extracted_json.get("destinatario", {})
        cnpj_tomador = tomador.get("cnpj", "Nao encontrado")

        if cnpj_tomador == "Nao encontrado":
            try:
                num_endereco = tomador.get("Numero", "")
                cidade = tomador.get("Cidade", "")
                if num_endereco and cidade:
                    result = cnpj(num_endereco=num_endereco, cidade=cidade)
                    cnpj_tomador = result if result else "Não encontrado"
                    extracted_json["destinatario"]["cnpj"] = cnpj_tomador
                    logging.info(f"Consultando CNPJ com número: {num_endereco}, cidade: {cidade}")
                    if cnpj_tomador != "Nao encontrado":
                        logging.info(f"CNPJ atualizado: {cnpj_tomador}")
                        logging.info(f"JSON atualizado: {extracted_json}")
                    else:
                        logging.warning("CNPJ não encontrado no banco de dados.")
            except Exception as e:
                logging.error(f"Erro ao buscar CNPJ: {e}")
                extracted_json["destinatario"]["cnpj"] = "Erro na consulta"
        
        valor_liquido = extracted_json.get("valor_liquido", "Informacao ausente")
        if valor_liquido == "Informacao ausente":
            extracted_json["valor_liquido"] = extracted_json.get("valor_total", "Informacao ausente")

        return extracted_json