import locale
import logging
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

locale.setlocale(locale.LC_ALL, '')

def _determine_header(notas: list) -> str:
    """
    Determina o cabeçalho da primeira coluna conforme o modelo da nota.
    Retorna "Número NF-e" se o modelo for "NFE" e "Número NFC-e" caso contrário.
    """
    if notas:
        modelo = notas[0].get("modelo", "NFC-E")
        return "Número NF-e" if modelo.upper() == "NFE" else "Número NFC-e"
    return "Número NFC-e"

def export_to_pdf(report: dict, output_file: str) -> None:
    """
    Exporta o relatório para PDF contendo:
      - Um resumo com Total de Notas e Notas Transmitidas;
      - Uma tabela principal com os campos:
          Número NF-e/NFC-e, Chave, Valor, Status, Emissão, Autorização;
      - Para cada nota, o detalhamento dos produtos.
    """
    doc = SimpleDocTemplate(
        output_file,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles["Title"],
        textColor=colors.darkblue,
        alignment=1
    )
    heading_style = ParagraphStyle(
        'HeadingStyle',
        parent=styles["Heading2"],
        textColor=colors.darkred
    )
    normal_style = styles["Normal"]

    story = []
    
    title = Paragraph("Relatório de NFC-e", title_style)
    story.append(title)
    story.append(Spacer(1, 12))
    
    # Totais
    notas = report.get("notas", [])
    total_notas = len(notas)
    valor_total = sum(n.get("valor", 0) for n in notas)
    notas_transmitidas = [n for n in notas if (n.get("status") or "").lower() == "autorizada"]
    total_transmitidas = len(notas_transmitidas)
    valor_transmitidas = sum(n.get("valor", 0) for n in notas_transmitidas)
    
    summary_header = Paragraph("Resumo", heading_style)
    story.append(summary_header)
    story.append(Spacer(1, 6))
    summary_data = [
        ["", "Quantidade", "Valor"],
        ["Total de Notas", total_notas, locale.currency(valor_total or 0, grouping=True)],
        ["Notas Transmitidas", total_transmitidas, locale.currency(valor_transmitidas or 0, grouping=True)]
    ]
    summary_table = Table(summary_data, colWidths=[150, 100, 150])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 12))
    
    # Cabeçalho da tabela principal
    num_header = _determine_header(notas)
    main_header = Paragraph("Notas", heading_style)
    story.append(main_header)
    story.append(Spacer(1, 6))
    table_header = [num_header, "Chave", "Valor", "Status", "Emissão", "Autorização"]
    table_data = [table_header]
    for nota in notas:
        nNF = nota.get("nNF", "N/A")
        chave = nota.get("chNFe") or "N/A"
        valor = locale.currency(nota.get("valor") or 0, grouping=True)
        status = nota.get("status") or ""
        emissao = nota.get("emitida") or ""
        autorizacao = nota.get("autorizada") or ""
        table_data.append([nNF, chave, valor, status, emissao, autorizacao])
    
    notes_table = Table(table_data, repeatRows=1, hAlign="CENTER")
    notes_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    story.append(notes_table)
    story.append(Spacer(1, 12))
    
    # Produtos de cada nota
    for nota in notas:
        story.append(Spacer(1, 12))
        nota_header = Paragraph(
            f"Produtos da Nota {nota.get('nNF', 'N/A')} - Chave: {nota.get('chNFe','N/A')}",
            styles["Heading3"]
        )
        story.append(nota_header)
        produtos = nota.get("produtos", [])
        if produtos:
            prod_data = []
            prod_header = ["Nome", "Código", "CFOP", "Qtd", "V. Unit", "V. Total"]
            prod_data.append(prod_header)
            for prod in produtos:
                v_unit = locale.currency(prod.get("valor_unitario") or 0, grouping=True)
                v_total = locale.currency(prod.get("valor_total") or 0, grouping=True)
                prod_row = [
                    prod.get("nome", "N/A"),
                    prod.get("codigo", "N/A"),
                    prod.get("cfop", "N/A"),
                    f"{prod.get('quantidade', 0)} {prod.get('unidade', '')}",
                    v_unit,
                    v_total
                ]
                prod_data.append(prod_row)
            prod_table = Table(prod_data, repeatRows=1, hAlign="CENTER")
            prod_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
            ]))
            story.append(prod_table)
        else:
            story.append(Paragraph("Nenhum produto registrado.", normal_style))
    
    doc.build(story)

def export_to_txt(report: dict, output_file: str) -> None:
    """
    Exporta o relatório para um arquivo TXT contendo:
      - Um resumo com Total de Notas e Notas Transmitidas (quantidade e valor);
      - Uma tabela com os campos: Número NF-e/NFC-e, Chave, Valor, Status, Emissão, Autorização;
      - Para cada nota, o detalhamento dos produtos.
    """
    lines = []
    lines.append("RELATÓRIO DE NFC-e")
    lines.append("=" * 80)
    lines.append("")
    # Totais
    notas = report.get("notas", [])
    total_notas = len(notas)
    valor_total = sum(n.get("valor", 0) for n in notas)
    notas_transmitidas = [n for n in notas if (n.get("status") or "").lower() == "autorizada"]
    total_transmitidas = len(notas_transmitidas)
    valor_transmitidas = sum(n.get("valor", 0) for n in notas_transmitidas)
    
    lines.append("Resumo:")
    lines.append(f"  Total de Notas: {total_notas} | Valor Total: {locale.currency(valor_total or 0, grouping=True)}")
    lines.append(f"  Notas Transmitidas: {total_transmitidas} | Valor Transmitido: {locale.currency(valor_transmitidas or 0, grouping=True)}")
    lines.append("")
    
    if notas:
        modelo = notas[0].get("modelo", "NFC-E")
    else:
        modelo = "NFC-E"
    num_header = "Número NF-e" if modelo.upper() == "NFE" else "Número NFC-e"
    
    header = "{:<15} {:<40} {:>12} {:<15} {:<12} {:<12}".format(
        num_header, "Chave", "Valor", "Status", "Emissão", "Autorização"
    )
    lines.append(header)
    lines.append("-" * len(header))
    
    for nota in notas:
        nNF = nota.get("nNF", "N/A")
        chave = nota.get("chNFe") or "N/A"
        try:
            valor = locale.currency(nota.get("valor") or 0, grouping=True)
        except Exception:
            valor = f"R$ {(nota.get('valor') or 0):,.2f}"
        status = nota.get("status") or ""
        emissao = nota.get("emitida") or ""
        autorizacao = nota.get("autorizada") or ""
        line = "{:<15} {:<40} {:>12} {:<15} {:<12} {:<12}".format(
            nNF, chave, valor, status, emissao, autorizacao
        )
        lines.append(line)
        if nota.get("produtos"):
            lines.append("  Produtos:")
            prod_header = "    {:<30} {:<10} {:<8} {:<10} {:>10} {:>10}".format(
                "Nome", "Código", "CFOP", "Qtd", "V. Unit", "V. Total"
            )
            lines.append(prod_header)
            lines.append("    " + "-" * len(prod_header))
            for prod in nota.get("produtos", []):
                try:
                    v_unit = locale.currency(prod.get("valor_unitario") or 0, grouping=True)
                except Exception:
                    v_unit = f"R$ {(prod.get('valor_unitario') or 0):,.2f}"
                try:
                    v_total = locale.currency(prod.get("valor_total") or 0, grouping=True)
                except Exception:
                    v_total = f"R$ {(prod.get('valor_total') or 0):,.2f}"
                prod_line = "    {:<30} {:<10} {:<8} {:<10} {:>10} {:>10}".format(
                    prod.get("nome", "N/A")[:30],
                    prod.get("codigo", "N/A"),
                    prod.get("cfop", "N/A"),
                    f"{prod.get('quantidade', 0)} {prod.get('unidade', '')}",
                    v_unit,
                    v_total
                )
                lines.append(prod_line)
    
    text_report = "\n".join(lines)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(text_report)

def export_to_csv(report: dict, output_file: str) -> None:
    """
    Exporta o relatório para CSV utilizando pandas, com as colunas:
      Número NF-e/NFC-e, Chave, Valor, Status, Emissão, Autorização.
    """
    notas = report.get("notas", [])
    if notas:
        modelo = notas[0].get("modelo", "NFC-E")
    else:
        modelo = "NFC-E"
    num_header = "Número NF-e" if modelo.upper() == "NFE" else "Número NFC-e"
    df = pd.DataFrame(notas)
    df = df.rename(columns={
        "nNF": num_header,
        "chNFe": "Chave",
        "valor": "Valor",
        "status": "Status",
        "emitida": "Emissão",
        "autorizada": "Autorização"
    })
    df = df[[num_header, "Chave", "Valor", "Status", "Emissão", "Autorização"]]
    df.to_csv(output_file, sep=";", index=False, float_format="%.2f", encoding="utf-8")

def export_to_excel(report: dict, output_file: str) -> None:
    """
    Exporta o relatório para Excel utilizando pandas, com as colunas:
      Número NF-e/NFC-e, Chave, Valor, Status, Emissão, Autorização.
    """
    notas = report.get("notas", [])
    if notas:
        modelo = notas[0].get("modelo", "NFC-E")
    else:
        modelo = "NFC-E"
    num_header = "Número NF-e" if modelo.upper() == "NFE" else "Número NFC-e"
    df = pd.DataFrame(notas)
    df = df.rename(columns={
        "nNF": num_header,
        "chNFe": "Chave",
        "valor": "Valor",
        "status": "Status",
        "emitida": "Emissão",
        "autorizada": "Autorização"
    })
    df = df[[num_header, "Chave", "Valor", "Status", "Emissão", "Autorização"]]
    df.to_excel(output_file, sheet_name="Relatorio", index=False, float_format="%.2f")
