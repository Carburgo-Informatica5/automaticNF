import os
import random
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

# Função para gerar dados aleatórios de notas fiscais
def gerar_notas_fiscais_aleatorias(quantidade, setores):
    """
    Gera uma lista de notas fiscais de forma aleatória.
    """
    emitentes = ["Empresa A", "Empresa B", "Empresa C", "Empresa D", "Empresa E", "Empresa F", "Empresa G"]
    notas = []
    data_inicio = datetime(2025, 1, 1)
    
    for i in range(quantidade):
        setor = random.choice(setores)
        numero_nota = str(random.randint(100, 999))
        data_lancamento = data_inicio + timedelta(days=random.randint(0, 30))  # Gera uma data aleatória em Janeiro
        emitente = random.choice(emitentes)
        valor_nota = f"{random.randint(500, 5000):.2f}"  # Valor entre 500 e 5000
        
        nota = {
            "número da nota": numero_nota,
            "data de lançamento": data_lancamento.strftime("%d.%m.%Y"),
            "setor": setor,
            "emitente": emitente,
            "valor da nota": valor_nota
        }
        notas.append(nota)
    
    return notas

def agrupar_por_setores(notas):
    """
    Agrupa as notas fiscais por setor e organiza cada grupo por data de lançamento.
    """
    agrupamento = {}
    for nota in notas:
        setor = nota.get("setor", "Sem Setor")
        if setor not in agrupamento:
            agrupamento[setor] = []
        agrupamento[setor].append(nota)
    
    # Ordenando por data de lançamento
    for setor, notas_setor in agrupamento.items():
        agrupamento[setor] = sorted(notas_setor, key=lambda x: datetime.strptime(x["data de lançamento"], "%d.%m.%Y"))
    
    return agrupamento

def filtrar_por_mes(notas, ano, mes):
    """
    Filtra as notas fiscais pelo mês e ano especificados.
    """
    notas_filtradas = []
    for nota in notas:
        data_lancamento = datetime.strptime(nota["data de lançamento"], "%d.%m.%Y")
        if data_lancamento.year == ano and data_lancamento.month == mes:
            notas_filtradas.append(nota)
    return notas_filtradas

def gerar_relatorios_mensais_por_setor(notas, diretorio_saida):
    """
    Gera relatórios mensais de notas fiscais separados por setor em formato PDF.
    Organiza os relatórios por setor e ano.
    """
    # Pegar a data do sistema para o ano e mês
    data_atual = datetime.now()
    ano = data_atual.year
    mes = data_atual.month
    
    notas_filtradas = filtrar_por_mes(notas, ano, mes)
    
    if not notas_filtradas:
        print(f"Nenhuma nota fiscal encontrada para {mes}/{ano}.")
        return

    agrupamento_setores = agrupar_por_setores(notas_filtradas)

    for setor, notas_setor in agrupamento_setores.items():
        diretorio_setor = os.path.join(diretorio_saida, setor)
        if not os.path.exists(diretorio_setor):
            os.makedirs(diretorio_setor)

        diretorio_ano = os.path.join(diretorio_setor, str(ano))
        if not os.path.exists(diretorio_ano):
            os.makedirs(diretorio_ano)

        caminho_arquivo = os.path.join(diretorio_ano, f"relatorio_{setor}_{mes:02d}_{ano}.pdf")
        gerar_pdf_relatorio(setor, notas_setor, caminho_arquivo)

def gerar_pdf_relatorio(setor, notas_setor, caminho_arquivo):
    """
    Gera o PDF para o relatório de um setor.
    """
    c = canvas.Canvas(caminho_arquivo, pagesize=letter)
    largura, altura = letter

    # Definindo título com destaque
    c.setFont("Helvetica-Bold", 20)
    c.setFillColor(HexColor("#00000"))  # Azul escuro para o título
    c.drawString(100, altura - 50, f"Relatório de Notas Fiscais - Setor: {setor}")

    # Separando com uma linha sutil
    c.setStrokeColor(HexColor("#B0BEC5"))
    c.setLineWidth(1)
    c.line(50, altura - 60, largura - 50, altura - 60)

    # Cabeçalho da tabela com fundo claro e bordas
    c.setFillColor(HexColor("#E8EAF6"))
    c.rect(50, altura - 100, largura - 100, 20, fill=True)
    
    colunas = ["Número da Nota", "Data de Lançamento", "Emitente", "Valor da Nota"]
    c.setFont("Helvetica-Bold", 10)
    largura_coluna = (largura - 100) / len(colunas)  # Largura proporcional para cada coluna
    for i, coluna in enumerate(colunas):
        c.setFillColor(HexColor("#1A237E"))
        # Centralizando o texto
        c.drawCentredString(50 + (i * largura_coluna) + largura_coluna / 2, altura - 95, coluna)

    # Desenhando as linhas da tabela com bordas e dados
    c.setFont("Helvetica", 10)
    y = altura - 120
    for nota in notas_setor:
        c.setStrokeColor(HexColor("#B0BEC5"))
        c.setLineWidth(0.5)

        # Desenhando as células da tabela
        for i in range(4):
            c.rect(50 + i * largura_coluna, y, largura_coluna, 20, stroke=True, fill=False)

        # Inserindo dados nas células (dados centralizados)
        c.setFont("Helvetica", 10)
        c.drawCentredString(50 + largura_coluna / 2, y + 5, nota["número da nota"])
        c.drawCentredString(50 + 3 * largura_coluna / 2, y + 5, nota["data de lançamento"])
        c.drawCentredString(50 + 5 * largura_coluna / 2, y + 5, nota["emitente"])
        c.drawCentredString(50 + 7 * largura_coluna / 2, y + 5, nota["valor da nota"])

        y -= 20

        # Se o espaço da página for pequeno, cria nova página
        if y < 50:
            c.showPage()
            c.setFont("Helvetica-Bold", 20)
            c.setFillColor(HexColor("#00000"))
            c.drawString(100, altura - 50, f"Relatório de Notas Fiscais - Setor: {setor}")
            c.line(50, altura - 60, largura - 50, altura - 60)
            c.setFont("Helvetica-Bold", 10)
            y = altura - 100

    c.save()
    print(f"Relatório gerado para o setor '{setor}': {caminho_arquivo}")

# Exemplo de uso
def main():
    setores = ["RH", "TI", "Financeiro"]
    notas_fiscais = gerar_notas_fiscais_aleatorias(50, setores)  # Gera 50 notas fiscais de forma aleatória
    
    diretorio_saida = "relatorios_mensais"
    gerar_relatorios_mensais_por_setor(notas_fiscais, diretorio_saida)

if __name__ == "__main__":
    main()
