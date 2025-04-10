import os
import re
import json
import logging
from decimal import Decimal
from gemini_api import GeminiAPI

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def verificar_impostos(json_data):
    """
    Verifica se os impostos retidos batem com a diferença entre o valor total e o valor líquido.
    Se o valor total for igual ao valor líquido, não deve haver impostos retidos.
    Atualizado para trabalhar com a estrutura JSON retornada pela API.
    """
    try:
        # Acessa os valores usando as chaves corretas do JSON retornado
        valor_total_str = json_data.get("valor_total", "0")
        valor_liquido_str = json_data.get("valor_liquido", "0")
        iss_retido = json_data.get("ISS_retido", "Não")

        # Converte valores para Decimal para precisão nos cálculos
        valor_total = Decimal(str(valor_total_str).replace(",", "."))
        valor_liquido = Decimal(str(valor_liquido_str).replace(",", "."))

        # Extrai os valores dos impostos retidos
        impostos_data = json_data.get("impostos", {})
        impostos = {
            "PIS": Decimal(str(impostos_data.get("PIS", "0")).replace(",", ".")),
            "COFINS": Decimal(str(impostos_data.get("COFINS", "0")).replace(",", ".")),
            "INSS": Decimal(str(impostos_data.get("INSS", "0")).replace(",", ".")),
            "ISS_retido": Decimal(str(impostos_data.get("ISS_retido", "0")).replace(",", ".")),
            "IR": Decimal(str(impostos_data.get("IR", "0")).replace(",", ".")),
            "CSLL": Decimal(str(impostos_data.get("CSLL", "0")).replace(",", ".")),
        }

        total_impostos = sum(impostos.values())

        if valor_total == valor_liquido:
            if total_impostos == 0:
                logging.info("Nenhum imposto retido, valores batem corretamente.")
            else:
                logging.warning(f"Valor Total e Valor Líquido são iguais, mas há impostos retidos: {impostos}")

        else:
            diferenca = valor_total - valor_liquido
            if abs(diferenca - total_impostos) > Decimal("0.01"):
                logging.error(f"Diferença ({diferenca}) não bate com impostos ({total_impostos}). Verifique possíveis erros.")

            if iss_retido == "Sim" and impostos["ISS_retido"] == Decimal("0.00"):
                logging.warning("O Gemini indicou que há ISS Retido, mas o valor extraído foi 0.00.")

            logging.info(f"Impostos ({total_impostos}) correspondem à diferença ({diferenca}).")

    except Exception as e:
        logging.error(f"Erro ao verificar impostos: {str(e)}")

def main():
    # Caminho absoluto para a pasta 'anexos'
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_folder = os.path.join(base_dir, 'anexos')
    json_folder = os.path.join(base_dir, '..', 'NOTAS EM JSON')  # Corrigido o caminho relativo

    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(json_folder, exist_ok=True)

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        logging.error("API key for Gemini is not set. Please set the GEMINI_API_KEY environment variable.")
        return

    gemini_api = GeminiAPI(api_key)

    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            logging.info(f"Processando PDF: {pdf_path}")

            try:
                upload_response = gemini_api.upload_pdf(pdf_path)
                if upload_response and upload_response.get("success"):
                    file_id = upload_response["file_id"]

                    status_response = gemini_api.check_processing_status(file_id)
                    if status_response and status_response.get("state") == "ACTIVE":
                        info = gemini_api.extract_info(file_id)
                        
                        if info and isinstance(info, dict):  # Agora esperamos um dicionário
                            try:
                                # Verifica a consistência dos impostos
                                verificar_impostos(info)

                                json_filename = os.path.splitext(pdf_file)[0] + ".json"
                                json_path = os.path.join(json_folder, json_filename)
                                
                                with open(json_path, "w", encoding="utf-8") as f:
                                    json.dump(info, f, ensure_ascii=False, indent=4)
                                logging.info(f"Informações extraídas e salvas em: {json_path}")
                                
                            except Exception as e:
                                logging.error(f"Erro ao processar ou salvar JSON para {pdf_file}: {str(e)}")
                        else:
                            logging.error(f"A resposta da API não é um dicionário válido para {pdf_file}")
            except Exception as e:
                logging.error(f"Erro inesperado ao processar PDF {pdf_path}: {str(e)}")

if __name__ == "__main__":
    main()