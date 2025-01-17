import os
import json
import sys
import xml.etree.ElementTree as ET
import re

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

        namespaces = {"ns": "http://www.portalfiscal.inf.br/nfe"}

        nota_fiscal_data = {
            "emitente": {},
            "destinatario": {},
            "chave_acesso": {},
            "num_nota": {},
            "data_emi": {},
            "data_venc": {},
            "modelo": {},
            "valor_total": [],
            "pagamento_parcelado": [],
            "info_adicional": {},
        }

        # Extrair informações do emitente
        emitente = root.find(".//ns:emit", namespaces)
        if emitente is not None:
            nota_fiscal_data["emitente"] = {
                "cnpj": emitente.findtext("ns:CNPJ", namespaces=namespaces),
                "nome": emitente.findtext("ns:xNome", namespaces=namespaces),
            }

        # Extrair informações do destinatário
        destinatario = root.find(".//ns:dest", namespaces)
        if destinatario is not None:
            nota_fiscal_data["destinatario"] = {
                "cnpj": destinatario.findtext("ns:CNPJ", namespaces=namespaces),
                "nome": destinatario.findtext("ns:xNome", namespaces=namespaces),
            }

        # Extrair informações da nota
        num_nota = root.find(".//ns:ide", namespaces)
        if num_nota is not None:
            nota_fiscal_data["num_nota"] = {
                "numero_nota": num_nota.findtext("ns:nNF", namespaces=namespaces)
            }

            data_emi = num_nota.findtext("ns:dhEmi", namespaces=namespaces)
            if data_emi is not None:
                data_emi_format = data_emi[:10].replace("-", "")
                data_emi_format = data_emi_format[6:] + data_emi_format[4:6] + data_emi_format[:4]  # Inverte a data
                nota_fiscal_data["data_emi"] = {"data_emissao": data_emi_format}
            else:
                nota_fiscal_data["data_emi"] = {"data_emissao": None}

            modelo = num_nota.findtext("ns:mod", namespaces=namespaces)
            if modelo:
                nota_fiscal_data["modelo"] = {"modelo": modelo}

        # Extrair data de vencimento
        data_venc = root.find(".//ns:cobr/ns:dup/ns:dVenc", namespaces=namespaces)
        if data_venc is not None:
            data_venc_format = data_venc.text.replace("-", "")
            data_venc_format = data_venc_format[6:] + data_venc_format[4:6] + data_venc_format[:4]  # Inverte a data
            nota_fiscal_data["data_venc"] = {"data_venc": data_venc_format}

        # Extrair chave de acesso
        chaveAcesso = root.find(".//ns:infNFe", namespaces)
        if chaveAcesso is not None and chaveAcesso.get("Id") is not None:
            chave_completa = chaveAcesso.get("Id")
            nota_fiscal_data["chave_acesso"] = {"chave": chave_completa[3:]}
        else:
            nota_fiscal_data["chave_acesso"] = {"chave": None}
        
        total = root.find(".//ns:ICMSTot", namespaces)
        if total is not None:
            valor_total = total.findtext("ns:vNF", namespaces=namespaces)
            if valor_total:
                nota_fiscal_data["valor_total"].append({
                    "valor_total": valor_total
                })
        
        # Extrair parcelas
        parcelas = root.findall(".//ns:cobr/ns:dup", namespaces)
        for parcela in parcelas:
            numero_parcela = parcela.findtext("ns:nDup", namespaces=namespaces)
            data_vencimento = parcela.findtext("ns:dVenc", namespaces=namespaces)
            valor_parcela = parcela.findtext("ns:vDup", namespaces=namespaces)
            if numero_parcela and data_vencimento and valor_parcela:
                nota_fiscal_data["pagamento_parcelado"].append({
                    "numero_parcela": numero_parcela,
                    "data_vencimento": data_vencimento,
                    "valor_parcela": valor_parcela
                })
        
        # Extrair informações adicionais de data e dias
        inf_adic = root.find(".//ns:infAdic/ns:infCpl", namespaces)
        if inf_adic is not None:
            inf_adic_text = inf_adic.text
            data_adicional = re.search(r'\d{4}-\d{2}-\d{2}', inf_adic_text)
            dias_adicional = re.search(r'\d+ dias', inf_adic_text)
            if data_adicional:
                nota_fiscal_data["data_adicional"] = {"data": data_adicional.group()}
            if dias_adicional:
                nota_fiscal_data["dias_adicional"] = {"dias": dias_adicional.group()}

        return nota_fiscal_data

    except ET.ParseError as e:
        print(f"Erro ao parsear o arquivo XML: '{xml_file_path}': {e}")
        return None


def testar_processamento_local():
    # Configurações de teste
    pasta_xml_teste = "anexos"

    try:
        # Verifica se a pasta existe
        if os.path.exists(pasta_xml_teste):
            # Itera sobre todos os arquivos na pasta
            for xml_file in os.listdir(pasta_xml_teste):
                if xml_file.endswith(".xml"):
                    dados_nf = parse_nota_fiscal(os.path.join(pasta_xml_teste, xml_file))
                    if dados_nf:
                        salvar_dados_em_arquivo(dados_nf, xml_file)
                    else:
                        print(f"Erro ao processar o arquivo: {xml_file}")
        else:
            print(f"Pasta não encontrada: {pasta_xml_teste}")
    except Exception as e:
        print(f"Erro durante o teste: {e}")


if __name__ == "__main__":
    testar_processamento_local()