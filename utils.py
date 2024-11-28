def format_currency(value):
    """
    Formata um valor numérico como moeda brasileira.

    Args:
        value (float): Valor numérico a ser formatado.

    Returns:
        str: Valor formatado como moeda (ex: R$ 1.234,56).
    """
    try:
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        raise ValueError(f"Erro ao formatar o valor: {value}. Erro: {str(e)}")


def validate_file_extension(file_path, valid_extensions):
    """
    Valida se o arquivo tem uma extensão permitida.

    Args:
        file_path (str): Caminho para o arquivo.
        valid_extensions (list): Lista de extensões válidas (ex: ['xml', 'zip']).

    Returns:
        bool: Verdadeiro se a extensão for válida, Falso caso contrário.
    """
    if not file_path:
        return False

    file_extension = file_path.split('.')[-1].lower()
    return file_extension in valid_extensions


def truncate_text(text, max_length=100):
    """
    Trunca um texto para o comprimento máximo especificado.

    Args:
        text (str): Texto a ser truncado.
        max_length (int): Comprimento máximo permitido.

    Returns:
        str: Texto truncado com "..." se exceder o comprimento máximo.
    """
    if len(text) > max_length:
        return text[:max_length - 3] + "..."
    return text


def parse_date(date_str, format="%Y-%m-%d"):
    """
    Converte uma string de data para um objeto de data.

    Args:
        date_str (str): Data como string (ex: "2023-11-21").
        format (str): Formato esperado da data (padrão: "%Y-%m-%d").

    Returns:
        datetime.date: Objeto de data correspondente ou None se inválido.
    """
    from datetime import datetime
    try:
        return datetime.strptime(date_str, format).date()
    except ValueError:
        return None


def generate_summary(notas):
    """
    Gera um resumo estatístico a partir de notas processadas.

    Args:
        notas (dict): Dicionário contendo os detalhes das NFC-e.

    Returns:
        dict: Resumo com total de notas, valor total, média, mínimo e máximo.
    """
    if not notas:
        return {
            "total_notes": 0,
            "total_value": 0.0,
            "avg_value": 0.0,
            "max_value": 0.0,
            "min_value": 0.0,
        }

    values = [nota['valor'] for nota in notas.values()]
    total_value = sum(values)
    total_notes = len(values)
    max_value = max(values, default=0.0)
    min_value = min(values, default=0.0)
    avg_value = total_value / total_notes if total_notes > 0 else 0.0

    return {
        "total_notes": total_notes,
        "total_value": total_value,
        "avg_value": avg_value,
        "max_value": max_value,
        "min_value": min_value,
    }


def pretty_summary(summary):
    """
    Converte o resumo gerado em uma string formatada.

    Args:
        summary (dict): Resumo gerado pela função generate_summary().

    Returns:
        str: String formatada para exibição do resumo.
    """
    return (
        f"Resumo da Análise\n"
        f"Total de Notas: {summary['total_notes']}\n"
        f"Valor Total: {format_currency(summary['total_value'])}\n"
        f"Valor Médio: {format_currency(summary['avg_value'])}\n"
        f"Valor Máximo: {format_currency(summary['max_value'])}\n"
        f"Valor Mínimo: {format_currency(summary['min_value'])}\n"
        f"{'-' * 50}\n"
    )
