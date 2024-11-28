# export.py
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import csv


def export_to_pdf(report_text, output_file):
    """
    Exporta o relatório como um arquivo PDF.

    Args:
        report_text (str): Texto do relatório a ser exportado.
        output_file (str): Caminho para o arquivo PDF de saída.
    """
    try:
        c = canvas.Canvas(output_file, pagesize=letter)
        width, height = letter
        y = height - 50  # Início do texto na parte superior

        # Configurações de fonte
        c.setFont("Helvetica", 12)

        # Escreve o texto linha por linha no PDF
        for line in report_text.split("\n"):
            if y < 50:  # Caso o espaço vertical termine, adiciona uma nova página
                c.showPage()
                c.setFont("Helvetica", 12)
                y = height - 50
            c.drawString(50, y, line)
            y -= 15  # Espaçamento entre linhas

        c.save()
    except Exception as e:
        raise RuntimeError(f"Erro ao exportar para PDF: {str(e)}")


def export_to_csv(report_text, output_file):
    """
    Exporta o relatório como um arquivo CSV.

    Args:
        report_text (str): Texto do relatório a ser exportado.
        output_file (str): Caminho para o arquivo CSV de saída.
    """
    try:
        # Estrutura para armazenar os dados
        dados = []
        nota_atual = {}
        produtos = []

        for line in report_text.split("\n"):
            if line.startswith("Resumo:"):
                continue  # Ignora o resumo
            elif line.startswith("Total de Notas:"):
                continue  # Ignora o total de notas
            elif line.startswith("Valor Total:"):
                continue  # Ignora o valor total
            elif line.startswith("Detalhes das Notas:"):
                continue  # Ignora o título
            elif line.startswith("Número NFC-e:"):
                if nota_atual:
                    nota_atual["produtos"] = produtos
                    dados.append(nota_atual)
                    nota_atual = {}
                    produtos = []
                nota_atual["Número NFC-e"] = line.split(":", 1)[1].strip()
            # elif line.startswith("Código Numérico (cNF):"):
                # nota_atual["Código Numérico (cNF)"] = line.split(":", 1)[1].strip()  # Removido
            elif line.startswith("Nome da Nota:"):
                nota_atual["Nome da Nota"] = line.split(":", 1)[1].strip()
            elif line.startswith("Valor:"):
                nota_atual["Valor (R$)"] = line.split(":", 1)[1].strip()
            elif line.startswith("Status:"):
                nota_atual["Status"] = line.split(":", 1)[1].strip()
            elif line.startswith("Data de Emissão:"):
                nota_atual["Data de Emissão"] = line.split(":", 1)[1].strip()
            elif line.startswith("Data de Autorização:"):
                nota_atual["Data de Autorização"] = line.split(":", 1)[1].strip()
            elif line.startswith("Produtos:"):
                continue  # Título dos produtos
            elif line.startswith("  - Nome:"):
                produto = {}
                partes = line.split(", ")
                for parte in partes:
                    chave, valor = parte.split(":", 1)
                    chave = chave.strip().replace("  - ", "")
                    valor = valor.strip()
                    produto[chave] = valor
                produtos.append(produto)
            elif line.strip() == "":
                continue  # Ignora linhas vazias
            else:
                continue  # Ignora outras linhas

        # Adiciona a última nota
        if nota_atual:
            nota_atual["produtos"] = produtos
            dados.append(nota_atual)

        # Define os cabeçalhos CSV
        cabeçalhos = ["Número NFC-e", "Nome da Nota", "Valor (R$)", "Status", "Data de Emissão", "Data de Autorização", "Produtos"]

        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=cabeçalhos)
            writer.writeheader()
            for nota in dados:
                # Converte a lista de produtos em uma string formatada
                produtos_str = "; ".join([f"{p.get('Nome', 'N/A')} (Código: {p.get('Código', 'N/A')}, CFOP: {p.get('CFOP', 'N/A')}, Quantidade: {p.get('Quantidade', 'N/A')}, Valor Unitário: {p.get('Valor Unitário', 'N/A')}, Valor Total: {p.get('Valor Total', 'N/A')})" for p in nota.get("produtos", [])])
                nota["Produtos"] = produtos_str
                writer.writerow(nota)
    except Exception as e:
        raise RuntimeError(f"Erro ao exportar para CSV: {str(e)}")


def export_to_txt(report_text, output_file):
    """
    Exporta o relatório como um arquivo TXT.

    Args:
        report_text (str): Texto do relatório a ser exportado.
        output_file (str): Caminho para o arquivo TXT de saída.
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as txtfile:
            txtfile.write(report_text)
    except Exception as e:
        raise RuntimeError(f"Erro ao exportar para TXT: {str(e)}")
