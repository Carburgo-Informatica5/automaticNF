import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def agrupar_por_setores(notas):
    """
    Agrupa as notas fiscais por setor.

    Args:
        notas (list): Lista de dicionários contendo os dados das notas fiscais.

    Returns:
        dict: Dicionário onde a chave é o setor e o valor é uma lista de notas desse setor.
    """
    agrupamento = {}
    for nota in notas:
        setor = nota.get("setor", "Sem Setor")  # Use "Sem Setor" como padrão se o setor não estiver definido
        if setor not in agrupamento:
            agrupamento[setor] = []
        agrupamento[setor].append(nota)
    return agrupamento

def filtrar_por_mes(notas, ano, mes):
    """
    Filtra as notas fiscais pelo mês e ano especificados, usando a data de lançamento.

    Args:
        notas (list): Lista de dicionários contendo os dados das notas fiscais.
        ano (int): Ano para filtrar.
        mes (int): Mês para filtrar.

    Returns:
        list: Lista de notas fiscais filtradas pelo mês e ano de lançamento.
    """
    notas_filtradas = []
    for nota in notas:
        # Usar a data de lançamento da nota
        data_lancamento = datetime.strptime(nota["data de lançamento"], "%Y-%m-%d")
        if data_lancamento.year == ano and data_lancamento.month == mes:
            notas_filtradas.append(nota)
    return notas_filtradas

def gerar_relatorios_mensais_por_setor(notas, ano, mes, diretorio_saida):
    """
    Gera relatórios mensais de notas fiscais separados por setor em formato PDF.
    (Agora organiza os relatórios por setor e dentro do ano)

    Args:
        notas (list): Lista de dicionários contendo os dados das notas fiscais.
        ano (int): Ano do relatório.
        mes (int): Mês do relatório.
        diretorio_saida (str): Caminho do diretório onde os relatórios serão salvos.
    """
    # Filtrar notas pelo mês e ano usando a data de lançamento
    notas_filtradas = filtrar_por_mes(notas, ano, mes)
    
    if not notas_filtradas:
        print(f"Nenhuma nota fiscal encontrada para {mes}/{ano}.")
        return

    # Agrupar notas por setor
    agrupamento_setores = agrupar_por_setores(notas_filtradas)

    # Gerar relatório para cada setor
    for setor, notas_setor in agrupamento_setores.items():
        # Criar diretório do setor, dentro do diretório do ano
        diretorio_setor = os.path.join(diretorio_saida, setor)
        if not os.path.exists(diretorio_setor):
            os.makedirs(diretorio_setor)

        # Criar diretório do ano dentro do setor
        diretorio_ano = os.path.join(diretorio_setor, str(ano))
        if not os.path.exists(diretorio_ano):
            os.makedirs(diretorio_ano)
        
        # Caminho do arquivo PDF para o relatório
        caminho_arquivo = os.path.join(diretorio_ano, f"relatorio_{setor}_{mes:02d}_{ano}.pdf")
        gerar_pdf_relatorio(setor, notas_setor, caminho_arquivo)

def gerar_pdf_relatorio(setor, notas_setor, caminho_arquivo):
    """
    Gera o PDF para o relatório de um setor.

    Args:
        setor (str): O nome do setor.
        notas_setor (list): Lista de notas fiscais do setor.
        caminho_arquivo (str): Caminho do arquivo PDF a ser gerado.
    """
    c = canvas.Canvas(caminho_arquivo, pagesize=letter)
    largura, altura = letter

    # Definindo título
    c.setFont("Helvetica-Bold", 16)
    c.drawString(100, altura - 50, f"Relatório de Notas Fiscais - Setor: {setor}")
    
    # Definindo a tabela
    c.setFont("Helvetica", 10)
    x = 50
    y = altura - 100
    colunas = ["Número da Nota", "Data de Lançamento", "Emitente", "Valor da Nota"]
    
    # Desenhando os cabeçalhos
    for i, coluna in enumerate(colunas):
        c.drawString(x + i * 120, y, coluna)

    # Desenhando as linhas da tabela
    y -= 20
    for nota in notas_setor:
        c.drawString(x, y, nota["número da nota"])  # Usando a chave 'número da nota'
        c.drawString(x + 120, y, nota["data de lançamento"])  # Usando a chave 'data de lançamento'
        c.drawString(x + 240, y, nota["emitente"])  # Usando a chave 'emitente'
        c.drawString(x + 360, y, nota["valor da nota"])  # Usando a chave 'valor da nota'
        y -= 20
    
    c.save()
    print(f"Relatório gerado para o setor '{setor}': {caminho_arquivo}")

# Exemplo de uso
def main():
    # Dados simulados de notas fiscais com as chaves conforme você deseja
    notas_fiscais = [
        {"número da nota": "123", "data de lançamento": "2025-01-01", "setor": "Financeiro", "emitente": "Empresa A", "valor da nota": "1000.00"},
        {"número da nota": "124", "data de lançamento": "2025-01-02", "setor": "RH", "emitente": "Empresa B", "valor da nota": "2000.00"},
        {"número da nota": "125", "data de lançamento": "2025-01-03", "setor": "Financeiro", "emitente": "Empresa C", "valor da nota": "1500.00"},
        {"número da nota": "126", "data de lançamento": "2025-01-15", "setor": "Logística", "emitente": "Empresa D", "valor da nota": "500.00"},
        {"número da nota": "127", "data de lançamento": "2025-01-20", "setor": "RH", "emitente": "Empresa E", "valor da nota": "3000.00"},
    ]

    # Configurar ano, mês e diretório de saída
    ano = 2025
    mes = 1
    diretorio_saida = "relatorios_mensais"

    # Gerar relatórios
    gerar_relatorios_mensais_por_setor(notas_fiscais, ano, mes, diretorio_saida)

if __name__ == "__main__":
    main()