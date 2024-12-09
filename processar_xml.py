import os
import json
import sys
import xml.etree.ElementTree as ET

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def salvar_dados_em_arquivo(dados_nf, nome_arquivo):
    """Salva os dados da nota fiscal em um arquivo JSON."""
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        json.dump(dados_nf, f, ensure_ascii=False, indent=4)
    print(f"Dados salvos em: {nome_arquivo}")

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

                    # Salva os dados extraídos em um novo arquivo
                    salvar_dados_em_arquivo(
                        dados_nf, f"dados_extraidos_{xml_file}.json"
                    )
        else:
            print(
                f"Pasta de teste não encontrada: {pasta_xml_teste}"
            )  # Mensagem corrigida
    except Exception as e:
        # Tratamento de erros genérico
        print("Erro durante o teste:")


if __name__ == "__main__":
    testar_processamento_local()
