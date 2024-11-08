import sys
import os

# Adiciona o diretório do projeto ao path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from programa.tratamentoErros import SimuladorBravos, tratadorErros
from programa.main import parse_nota_fiscal, processar_notas_fiscais

def testar_processamento_local():
    # Configurações de teste
    pasta_xml_teste = "dados_teste/xml_mock"  # Crie esta pasta
    
    # Inicializa simuladores
    simulador = SimuladorBravos(modo_teste=True)
    tratador_erros = tratadorErros()
    
    try:
        # Teste de conexão
        conexao_sucesso = simulador.conectar()
        print("Conexão simulada:", "Sucesso" if conexao_sucesso else "Falha")
        
        # Teste de parsing de XML
        xml_teste = os.path.join(pasta_xml_teste, "nota_fiscal_1.xml")
        
        if os.path.exists(xml_teste):
            dados_nf = parse_nota_fiscal(xml_teste)
            print("Dados extraídos:", dados_nf)
        else:
            print(f"Arquivo de teste não encontrado: {xml_teste}")
        
        # Processar notas fiscais em modo de teste
        processar_notas_fiscais(pasta_xml_teste)
        
    except Exception as e:
        # Tratamento de erros genérico
        info_erro = tratador_erros.tratarErros(e, "Teste Local")
        print("Erro durante o teste:", info_erro)

if __name__ == "__main__":
    testar_processamento_local()