# processing.py
import os
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET


def analyze_file(file_path, progress_dialog=None):
    """
    Processa e analisa um arquivo XML ou ZIP.
    
    Args:
        file_path (str): Caminho para o arquivo XML ou ZIP.
        progress_dialog (QProgressDialog, opcional): Diálogo de progresso para atualizações.

    Returns:
        dict: Um dicionário contendo os detalhes das NFC-e processadas.
    """
    temp_dir = tempfile.mkdtemp()
    try:
        xml_files = extract_files([file_path], temp_dir)
        return process_xml_files(xml_files, progress_dialog)
    finally:
        shutil.rmtree(temp_dir)


def extract_files(files, destination):
    """
    Extrai arquivos XML de um ou mais arquivos ZIP ou pastas.

    Args:
        files (list): Lista de caminhos para arquivos ou pastas.
        destination (str): Caminho para o diretório onde os arquivos XML serão extraídos.
    
    Returns:
        list: Lista de caminhos para os arquivos XML extraídos.
    """
    xml_files = []
    for file in files:
        if os.path.isfile(file):
            if zipfile.is_zipfile(file):
                with zipfile.ZipFile(file, 'r') as zip_ref:
                    zip_ref.extractall(destination)
            else:
                shutil.copy(file, destination)
        elif os.path.isdir(file):
            for root, _, filenames in os.walk(file):
                for filename in filenames:
                    if filename.endswith('.xml'):
                        full_path = os.path.join(root, filename)
                        shutil.copy(full_path, destination)
    extracted_files = [os.path.join(destination, f) for f in os.listdir(destination) if f.endswith('.xml')]
    print(f"Arquivos XML extraídos: {extracted_files}")  # Log para depuração
    return extracted_files


def process_xml_files(xml_files, progress_dialog=None):
    """
    Processa e extrai informações de uma lista de arquivos XML.

    Args:
        xml_files (list): Lista de caminhos para arquivos XML.
        progress_dialog (QProgressDialog, opcional): Diálogo de progresso para atualizações.
    
    Returns:
        dict: Um dicionário contendo os detalhes de cada NFC-e processada.
    """
    notas = []
    total_files = len(xml_files)

    for i, xml_file in enumerate(xml_files):
        try:
            nota_details = extract_note_details(xml_file)
            if nota_details:
                notas.append(nota_details)
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {e}")

        # Atualizar progresso
        if progress_dialog:
            progress_value = int((i + 1) / total_files * 100)
            progress_dialog.setValue(progress_value)
            if progress_dialog.wasCanceled():
                print("Processo de análise cancelado pelo usuário.")
                break

    # Resumo geral
    resumo = {
        "total_notas": len(notas),
        "valor_total": sum(nota["valor"] for nota in notas if "valor" in nota),
    }

    return {
        "resumo": resumo,
        "notas": notas,
    }


def extract_note_details(xml_file):
    """
    Extrai detalhes de uma nota fiscal eletrônica de um arquivo XML.

    Args:
        xml_file (str): Caminho para o arquivo XML.

    Returns:
        dict: Um dicionário contendo os detalhes da nota fiscal eletrônica.
    """
    detalhes = {
        "nome": os.path.basename(xml_file),
        "nNF": "N/A",  # Número da NFC-e
        "cNF": "N/A",   # Código Numérico da NFC-e
        "valor": 0.0,
        "status": "Desconhecido",
        "codigo_status": "Não disponível",
        "autorizada": None,
        "emitida": None,
        "cancelada": False,
        "produtos": [],
        "emitente": {},
    }
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Namespace padrão
        namespace = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

        # Número da NFC-e (nNF) e Código Numérico (cNF)
        nNF = root.find(".//nfe:ide/nfe:nNF", namespace)
        if nNF is not None:
            detalhes["nNF"] = nNF.text

        cNF = root.find(".//nfe:ide/nfe:cNF", namespace)
        if cNF is not None:
            detalhes["cNF"] = cNF.text

        # Valor total da nota
        valor_total = root.find(".//nfe:vNF", namespace)
        if valor_total is not None:
            detalhes["valor"] = float(valor_total.text)

        # Código e status da nota
        cStat = root.find(".//nfe:cStat", namespace)
        if cStat is not None:
            detalhes["codigo_status"] = cStat.text
            if cStat.text == '100':  # Nota autorizada
                detalhes["status"] = "Autorizada"
            elif cStat.text in ['101', '135', '151']:  # Nota cancelada
                detalhes["status"] = "Cancelada"
                detalhes["cancelada"] = True
            else:
                detalhes["status"] = "Desconhecido"

        # Datas de autorização e emissão
        dhRecbto = root.find(".//nfe:dhRecbto", namespace)
        if dhRecbto is not None:
            detalhes["autorizada"] = dhRecbto.text[:10]  # Data em formato yyyy-MM-dd

        dhEmi = root.find(".//nfe:dhEmi", namespace)
        if dhEmi is not None:
            detalhes["emitida"] = dhEmi.text[:10]  # Data em formato yyyy-MM-dd

        # Emitente
        emitente = root.find(".//nfe:emit", namespace)
        if emitente is not None:
            detalhes["emitente"] = {
                "nome": emitente.find("nfe:xNome", namespace).text if emitente.find("nfe:xNome", namespace) is not None else "Desconhecido",
                "cnpj": emitente.find("nfe:CNPJ", namespace).text if emitente.find("nfe:CNPJ", namespace) is not None else "N/A",
                "endereco": f"{emitente.find('nfe:xLgr', namespace).text if emitente.find('nfe:xLgr', namespace) is not None else ''}, "
                            f"{emitente.find('nfe:nro', namespace).text if emitente.find('nfe:nro', namespace) is not None else ''}, "
                            f"{emitente.find('nfe:xBairro', namespace).text if emitente.find('nfe:xBairro', namespace) is not None else ''}, "
                            f"{emitente.find('nfe:xMun', namespace).text if emitente.find('nfe:xMun', namespace) is not None else ''} - "
                            f"{emitente.find('nfe:UF', namespace).text if emitente.find('nfe:UF', namespace) is not None else ''}"
            }

        # Produtos da nota
        produtos = []
        for det in root.findall(".//nfe:det", namespace):
            prod = det.find("nfe:prod", namespace)
            if prod is not None:
                produto = {
                    "nome": prod.find("nfe:xProd", namespace).text if prod.find("nfe:xProd", namespace) is not None else "Desconhecido",
                    "codigo": prod.find("nfe:cProd", namespace).text if prod.find("nfe:cProd", namespace) is not None else "N/A",
                    "cfop": prod.find("nfe:CFOP", namespace).text if prod.find("nfe:CFOP", namespace) is not None else "N/A",  # Extração do CFOP
                    "quantidade": float(prod.find("nfe:qCom", namespace).text) if prod.find("nfe:qCom", namespace) is not None else 0.0,
                    "unidade": prod.find("nfe:uCom", namespace).text if prod.find("nfe:uCom", namespace) is not None else "",
                    "valor_unitario": float(prod.find("nfe:vUnCom", namespace).text) if prod.find("nfe:vUnCom", namespace) is not None else 0.0,
                    "valor_total": float(prod.find("nfe:vProd", namespace).text) if prod.find("nfe:vProd", namespace) is not None else 0.0,
                }
                produtos.append(produto)
                print(f"Produto extraído: {produto}")  # Log para depuração

        detalhes["produtos"] = produtos

    except ET.ParseError:
        detalhes["status"] = "Erro ao processar XML"
    except Exception as e:
        detalhes["status"] = f"Erro: {str(e)}"
    
    return detalhes
