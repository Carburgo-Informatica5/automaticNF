import os
import time
import google.generativeai as genai

class GeminiAPI:
    def __init__(self, api_key):
        genai.configure(api_key=api_key)

    def upload_pdf(self, path, mime_type="application/pdf"):
        """Uploads the given PDF file to Gemini."""
        file = genai.upload_file(path, mime_type=mime_type)
        print(f"Uploaded file '{file.display_name}' as: {file.uri}")
        return {"success": True, "file_id": file.name}

    def check_processing_status(self, file_id):
        """Waits for the given files to be active."""
        print("Waiting for file processing...")
        file = genai.get_file(file_id)
        while file.state.name == "PROCESSING":
            print(".", end="", flush=True)
            time.sleep(10)
            file = genai.get_file(file_id)
        if file.state.name != "ACTIVE":
            raise Exception(f"File {file.name} failed to process")
        print("...file ready")
        return {"state": file.state.name}

    def extract_info(self, file_id):
        """Extrai informações do PDF usando Gemini."""
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

        # Obtém o objeto do arquivo carregado
        file_obj = genai.get_file(file_id)

        # Cria uma sessão de chat vazia
        chat_session = model.start_chat(history=[])

        # Envia a mensagem inicial com referência ao arquivo
        response = chat_session.send_message([
            {"text": "Extraia as seguintes informações da nota fiscal em formato JSON:\n"
                    "Numero da nota\n"
                    "Data da emissão (formato DDMMYYYY, sem barras)\n"
                    "Nome do prestador de serviço\n"
                    "Nome do tomador do serviço\n"
                    "CNPJ do prestador de serviço (somente números, sem pontos ou traços) se não extrair pare o processo\n"
                    "CNPJ do tomador do serviço (somente números, sem pontos ou traços)  se não extrair pare o processo\n"
                    "Valor total (substituir vírgulas por pontos)\n"
                    "Valor líquido (substituir vírgulas por pontos)\n"
                    "Verifique se há ISS retido:\n"
                    "- Se houver ISS retido, adicione o campo \"ISS Retido\": \"Sim\".\n"
                    "- Caso contrário, adicione o campo \"ISS Retido\": \"Não\".\n"
                    "Extraia os valores dos seguintes impostos retidos, se houver:\n"
                    "- PIS\n"
                    "- COFINS\n"
                    "- INSS\n"
                    "- ISS Retido (valor real, não o ISS total)\n"
                    "- IR\n"
                    "- CSLL\n"
                    "Se algum imposto não estiver presente, retorne o valor '0.00'.\n"
                    "Certifique-se de que o JSON esteja formatado corretamente e sem caracteres fictícios e com assentos.\n"},
            file_obj
            
        ])

        return response.text