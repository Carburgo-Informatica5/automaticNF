import os
import json
import sys
from programa.main import parse_nota_fiscal
from programa.tratamentoErros import tratadorErros

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def salvar_dados_em_arquivo(dados_nf, nome_arquivo):
    """Salva os dados da nota fiscal em um arquivo JSON."""
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        json.dump(dados_nf, f, ensure_ascii=False, indent=4)
    print(f"Dados salvos em: {nome_arquivo}")

def testar_processamento_local():
    # Configurações de teste
    pasta_xml_teste = "dados_teste"
    
    # Inicializa o tratador de erros
    tratador_erros = tratadorErros()
    
    try:
        # Verifica se a pasta existe
        if os.path.exists(pasta_xml_teste):
            # Itera sobre todos os arquivos na pasta
            for xml_file in os.listdir(pasta_xml_teste):
                if xml_file.endswith('.xml'):  # Verifica se o arquivo é um XML
                    xml_teste = os.path.join(pasta_xml_teste, xml_file)
                    dados_nf = parse_nota_fiscal(xml_teste)
                    print("Dados extraídos:", dados_nf)
                    
                    # Salva os dados extraídos em um novo arquivo
                    salvar_dados_em_arquivo(dados_nf, f"dados_extraidos_{xml_file}.json")
        else:
            print(f"Pasta de teste não encontrada: {pasta_xml_teste}")  # Mensagem corrigida
    except Exception as e:
        # Tratamento de erros genérico
        info_erro = tratador_erros.tratarErros(e, "Teste Local")
        print("Erro durante o teste:", info_erro)

if __name__ == "__main__":
    testar_processamento_local()