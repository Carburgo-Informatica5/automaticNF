import os
import json
import sys
import xml.etree.ElementTree as ET

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def salvar_dados_em_arquivo(dados_nf, nome_arquivo, pasta_destino="NOTA EM JSON"):
    """Salva os dados da nota fiscal em um arquivo JSON."""
    nome_arquivo = os.path.splitext(nome_arquivo)[0]
    nome_arquivo = f"dados_nota_{nome_arquivo}.json"
    caminho_completo = os.path.join(pasta_destino, nome_arquivo)
    os.makedirs(pasta_destino, exist_ok=True)
    with open(caminho_completo, "w", encoding="utf-8") as f:
        json.dump(dados_nf, f, ensure_ascii=False, indent=4)
    print(f"Dados salvos em: {caminho_completo}")


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
            "eminente": {},
            "destinatario": {},
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
            nota_fiscal_data["eminente"] = {
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
        if forma_pagamento:
            for pagamento in forma_pagamento:
                parcelas = pagamento.findall("ns0:dup", namespaces=namespaces)
                if parcelas:
                    # Se existem parcelas, itere sobre elas
                    for parcela in parcelas:
                        nmr_parc = parcela.findtext("ns0:nDup", namespaces=namespaces)
                        data_venc = parcela.findtext("ns0:dVenc", namespaces=namespaces)
                        valor_parc = parcela.findtext("ns0:vDup", namespaces=namespaces)

                        if nmr_parc and data_venc and valor_parc:
                            data_venc_format = data_venc[:10].replace("-", " ")
                            nota_fiscal_data["pagamento_parcelado"].append(
                                {
                                    "nmr_parc": nmr_parc,
                                    "data_venc": f"{data_venc_format[8:10]}{data_venc_format[5:7]}{data_venc_format[0:4]}",
                                    "valor_parc": valor_parc,
                                }
                            )
                else:
                    # Se não existem parcelas, verificar se há um pagamento único
                    valor_total = pagamento.findtext("ns0:fat/ns0:vLiq", namespaces=namespaces)
                    data_venc = pagamento.findtext("ns0:dup/ns0:dVenc", namespaces=namespaces)

                    if valor_total and data_venc:
                        data_venc_format = data_venc[:10].replace("-", " ")
                        # Tratar como pagamento único
                        nota_fiscal_data["valor_total"].append(
                            {
                                "valor_total": valor_total,
                                "data_venc": f"{data_venc_format[8:10]}{data_venc_format[5:7]}{data_venc_format[0:4]}"
                            }
                        )
        else:
            print("Forma de pagamento não encontrada")

        # Extrair data de emissão
        data_emi = num_nota.findtext("ns0:dhEmi", namespaces=namespaces)
        if data_emi:
            data_emi_format = data_emi[:10].replace("-", " ")
            nota_fiscal_data["data_emi"] = {
                "data_emissao": f"{data_emi_format[8:10]}{data_emi_format[5:7]}{data_emi_format[0:4]}"
            }
        else:
            print("Data de emissão não encontrada")

        # Extrair data de validade
        data_vali = num_nota.findtext("ns0:dhSaiEnt", namespaces=namespaces)
        if data_vali:
            data_vali_format = data_vali[:10].replace("-", " ")
            nota_fiscal_data["data_vali"] = {
                "data_validade": f"{data_vali_format[8:10]}{data_vali_format[5:7]}{data_vali_format[0:4]}"
            }
        else:
            print("Data de validade não encontrada")

        # Extrair modelo
        modelo = num_nota.findtext("ns0:mod", namespaces=namespaces)
        if modelo:
            nota_fiscal_data["modelo"] = {"modelo": modelo}
        else:
            print("Modelo não encontrado")

        # Extrair chave de acesso
        chaveAcesso = root.find(".//ns0:infNFe", namespaces)
        if chaveAcesso:
            chave_completa = chaveAcesso.get("Id")
            nota_fiscal_data["chave_acesso"] = {
                "chave": chave_completa[3:]  # Remove os 3 primeiros caracteres
            }
        else:
            print("Chave de acesso não encontrada")

        return nota_fiscal_data
    except ET.ParseError as e:
        print(f"Erro ao parsear o arquivo XML: '{xml_file_path}': {e}")
        return None
    except Exception as e:
        print(f"Erro ao extrair dados do XML: '{xml_file_path}': {e}")
        return None


def testar_processamento_local():
    # Configurações de teste
    pasta_xml_teste = "anexos"

    try:
        # Verifica se a pasta existe
        if os.path.exists(pasta_xml_teste):
            # Itera sobre todos os arquivos na pasta
            for xml_file in os.listdir(pasta_xml_teste):
                if xml_file.endswith(".xml"):  # Verifica se o arquivo é um XML
                    xml_teste = os.path.join(pasta_xml_teste, xml_file)
                    dados_nf = parse_nota_fiscal(xml_teste)
                    if dados_nf:
                        salvar_dados_em_arquivo(dados_nf, xml_file[:-4])
        else:
            print(f"Pasta não encontrada: {pasta_xml_teste}")
    except Exception as e:
        # Tratamento de erros genérico
        print(f"Erro durante o teste: {e}")


if __name__ == "__main__":
    testar_processamento_local()