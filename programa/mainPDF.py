data_vencimento = valores_extraidos[
    "data_vencimento"
]
tipo_imposto = valores_extraidos["tipo_imposto"]

if "json_path" in dados_nota_fiscal and dados_nota_fiscal["json_path"]:
    try:
        with open(dados_nota_fiscal["json_path"], "r") as f:
            json_data = json.load(f)
            logging.info(f"JSON carregado do PDF: {json.dumps(json_data, indent=4)}")
    except Exception as e:
        logging.error(f"Erro ao carregar JSON do PDF: {e}")
        json_data = None
else:
    json_data = None  # Evita erro caso json_path não exista
    # Se for um PDF, formatamos os dados usando json_data
    if json_data is not None:
        logging.info(f"Usando dados extraídos do PDF: {json.dumps(json_data, indent=4)}")
        dados_nota_fiscal_formatado = format_nota_fiscal(json_data, dados_email)
    else:
        logging.info("Nenhum JSON carregado, pois o anexo não é um PDF.")
        dados_nota_fiscal_formatado = format_nota_fiscal(dados_nota_fiscal, dados_email)
    if not dados_nota_fiscal_formatado:
        logging.warning("Aviso: Nenhum dado formatado foi encontrado! Usando valores padrão.")
    else:
        dados_nota_fiscal = dados_nota_fiscal_formatado
    # Mapeia os campos do JSON para o formato esperado
    dados_nota_fiscal = {
        "valor_total": [
            {
                "valor_total": json_data[
                    "valor_total"
                ]["valor_total"]
            }
        ],
        "emitente": {
            "nome": json_data["emitente"]["nome"],
            "cnpj": json_data["emitente"]["cnpj"],
        },
        "num_nota": {
            "numero_nota": json_data["num_nota"][
                "numero_nota"
            ]
        },
        "data_emi": {
            "data_emissao": json_data["data_emi"][
                "data_emissao"
            ]
        },
        "data_venc": {
            "data_venc": json_data.get("data_venc", {}).get("data_venc") or dados_email.get("data_venc_nfs")
        },
        "chave_acesso": {
            "chave": json_data["chave_acesso"][
                "chave"
            ]
        },
        "modelo": {
            "modelo": json_data["modelo"]["modelo"]
        },
        "destinatario": {
            "nome": json_data["destinatario"][
                "nome"
            ],
            "cnpj": json_data["destinatario"][
                "cnpj"
            ],
        },
        "pagamento_parcelado": [],
        "serie": "",  
    }
logging.info(f"Valor de json_data['data_venc']: {json_data.get('data_venc')}")
logging.info(f"Valor de dados_email['data_venc_nfs']: {dados_email.get('data_venc_nfs')}")
logging.info(f"Valor de data_venc_nfs ao criar dados_email: {data_vencimento}")

def clean_extracted_json(json_data):
    """Remove duplicatas e ajusta formatação do JSON extraído."""
    if "ISS Retido" in json_data:
        if (
            isinstance(json_data["ISS Retido"], str)
            and json_data["ISS Retido"].lower() == "não"
        ):
            json_data["ISS Retido"] = "Não"
        else:
            json_data["ISS Retido"] = f"{float(json_data['ISS Retido']):.2f}"
    return json_data

def map_json_fields(json_data, dados_email):
    
    logging.info(f"Data de vencimento mapeada: {data_venc}, {data_venc_nfs}")
    mapped_data = {
        "emitente": {
            "cnpj": json_data.get("CNPJ do prestador de serviço"),
            "nome": json_data.get("Nome do prestador de serviço"),
        },
        "destinatario": {
            "cnpj": json_data.get("CNPJ do tomador do serviço"),
            "nome": json_data.get("Nome do tomador do serviço"),
        },
        "num_nota": {
            "numero_nota": json_data.get("Numero da nota"),
        },
        "data_venc": {
            "data_venc": dados_email.get("data_venc_nfs"), 
        },
        "data_emi": {
            "data_emissao": json_data.get("Data da emissão"),
        },
        "valor_total": {
            "valor_total": json_data.get("Valor total"),
        },
        "valor_liquido": {
            "valor_liquido": json_data.get("Valor líquido"),
        },
        "modelo": {
            "modelo": "01",
        },
        "serie": {
            "serie": "1",
        },
        "chave_acesso": {
            "chave": "",
        },
        "impostos": {
            "ISS_retido": json_data.get("ISS retido"),
            "PIS": json_data.get("PIS"),
            "COFINS": json_data.get("COFINS"),
            "INSS": json_data.get("INSS"),
            "IR": json_data.get("IR"),
            "CSLL": json_data.get("CSLL"),
        },
    }
    return mapped_data


def process_pdf(pdf_path, dados_email):
    json_folder = os.path.abspath(
        os.path.join("C:/Users/VAS MTZ/Desktop/Caetano Apollo/NOTA EM JSON")
    )

    os.makedirs(json_folder, exist_ok=True)

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        logging.error("API key not found")
        return

    gemini_api = GeminiAPI(api_key)

    # Verifica se o PDF é um arquivo válido
    if not os.path.isfile(pdf_path) or not pdf_path.endswith(".pdf"):
        logging.error(f"O caminho {pdf_path} não é um arquivo PDF válido")
        return

    try:
        upload_response = gemini_api.upload_pdf(pdf_path)
        if not upload_response.get("success"):
            logging.error(f"Erro ao fazer upload do PDF: {pdf_path}")
            return

        file_id = upload_response["file_id"]

        status_response = gemini_api.check_processing_status(file_id)
        if status_response.get("state") != "ACTIVE":
            logging.error(f"Erro no processamento do arquivo {pdf_path}")
            return

        extracted_text = gemini_api.extract_info(file_id)
        if not extracted_text:
            logging.error(f"Erro ao extrair informações do PDF: {pdf_path}")
            return

        extracted_text = re.sub(
            r"^```json", "", extracted_text
        ).strip()  # Remove o ```json do início
        extracted_text = re.sub(
            r"```$", "", extracted_text
        ).strip()  # Remove o ``` do final

        extracted_json = json.loads(
            extracted_text
        )  # Converte string JSON para dicionário
        cleaned_json = clean_extracted_json(extracted_json)  # Limpa o JSON
        mapped_json = map_json_fields(cleaned_json, dados_email)  # Mapeia os campos do JSON

        json_path = os.path.join(
            json_folder, f"{os.path.splitext(os.path.basename(pdf_path))[0]}.json"
        )
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(
                mapped_json, f, ensure_ascii=False, indent=4
            )  # Salva como JSON formatado
        logging.info(f"JSON extraído salvo em: {json_path}")

    except json.JSONDecodeError:
        logging.error(
            f"Erro: A resposta da API não é um JSON válido após limpeza: {extracted_text}"
        )
    except Exception as e:
        logging.error(f"Erro ao processar PDF {pdf_path}: {e}")


    elif filename.lower().endswith(".pdf"):
        process_pdf(filepath, dados_email)
        # Retorna um dicionário com o caminho do JSON salvo
        json_filename = f"{os.path.splitext(os.path.basename(filepath))[0]}.json"
        json_path = os.path.join(
            "C:/Users/VAS MTZ/Desktop/Caetano Apollo/NOTA EM JSON", json_filename
        )
        return {"json_path": json_path}
    

if anexo.endswith(".pdf"):
                # Se arquivo endswitch (.pdf)
                gui.press("right", presses=1)
                gui.press("tab", presses=20)
                gui.write(cod_item)
                gui.press("tab", presses=10)
                gui.write("1")
                gui.press("tab")
                gui.write(valor_total)
                gui.press("tab", presses=34)
                gui.write(descricao)

                impostos = dados_email.get("impostos", {})

                def preencher_campo(valor, tabs):
                    """Preenche o campo apenas se o valor for diferente de '0.00' e não for None"""
                    if valor and valor != "0.00":
                        gui.press("tab", presses=tabs)
                        gui.write(valor)

                        # Sequência de preenchimento com verificação
                        gui.press("tab", presses=43)
                        gui.write(valor_total)

                        preencher_campo(impostos.get("PIS"), 1)
                        preencher_campo(valor_total, 5)

                        preencher_campo(impostos.get("COFINS"), 1)
                        preencher_campo(valor_total, 5)

                        preencher_campo(impostos.get("CSLL"), 1)
                        preencher_campo(valor_total, 9)

                        preencher_campo(impostos.get("ISS retido"), 7)

                #! Calculo de dias para impostos
                hoje = datetime.now()

                # Determina o próximo mês
                if hoje.month == 12:
                    proximo_mes = 1
                    ano = hoje.year + 1
                else:
                    proximo_mes = hoje.month + 1
                    ano = hoje.year

                # Calcula o primeiro dia do próximo mês
                primeiro_dia_proximo_mes = datetime(ano, proximo_mes, 1)

                # Determina o dia 20 do próximo mês
                dia_20_proximo_mes = datetime(ano, proximo_mes, 20)

                # Calcula a diferença em dias
                dias_restantes = (dia_20_proximo_mes - hoje).days

                gui.press("tab", presses=23)
                gui.press("left")
                gui.press("tab", presses=5)
                gui.press("enter")

                impostos = dados_email.get("impostos", {})

                # Função para somar apenas valores diferentes de "0.00" ou None
                def somar_impostos(*valores):
                    return sum(
                        float(valor) for valor in valores if valor and valor != "0.00"
                    )

                # Calcula a soma de PIS, COFINS e CSLL (se forem diferentes de 0.00)
                PCC = somar_impostos(
                    impostos.get("PIS"), impostos.get("COFINS"), impostos.get("CSLL")
                )

                # Exibe o valor calculado para depuração
                logging.info(f"Valor de PCC calculado: {PCC:.2f}")

                if INSS != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down")
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(INSS)
                    gui.press("tab", "enter")
                if IR != "0.00":
                    gui.press("tab", presses=7)
                    if tipo_imposto == "normal":
                        gui.press("down", presses=2)
                    elif tipo_imposto == "comissão":
                        gui.press("down", presses=7)
                    else:
                        gui.press("down", presses=5)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(IR)
                    gui.press("tab", "enter")
                if PCC != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down", presses=3)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(PCC)
                    gui.press("tab", "enter")
                if ISS_retido != "0.00":
                    gui.press("tab", presses=7)
                    gui.press("down", presses=4)
                    gui.press("tab")
                    gui.write(dias_restantes)
                    gui.press("tab", presses=2)
                    gui.write(ISS_retido)
                    gui.press("tab", "enter")
                gui.press("tab", "enter")
                gui.press("tab", presses=5)
                gui.write(data_venc_nfs)
                gui.press("tab", presses=4)
                gui.write(valor_liquido)
                gui.press("tab", presses=3)
                gui.press("enter")
                gui.press("tab", presses=39)
                gui.press("enter")
                pass
            
                pdf_path = os.path.join(DIRECTORY, "anexos")
                extracted_info = process_pdf(pdf_path, dados_email)
                if os.path.exists(pdf_path):
                    extracted_info = process_pdf(pdf_path, dados_email)
                else:
                    logging.error("Erro ao precessar o PDF")