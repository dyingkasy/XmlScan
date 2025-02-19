import os
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET
import datetime
import logging

logging.basicConfig(
    filename="app.log",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

def load_official_keys() -> set:
    filepath = "keys.csv"
    if not os.path.exists(filepath):
        return set()
    oficial = set()
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or "," not in line:
                continue
            parts = line.split(",")
            if len(parts) < 3:
                continue
            nNF, cNF, cnpj = parts[0].strip(), parts[1].strip(), parts[2].strip()
            oficial.add((nNF, cNF, cnpj))
    return oficial

def analyze_file(file_path: str, progress_dialog=None) -> dict:
    import_path = os.path.abspath(file_path)
    temp_dir = tempfile.mkdtemp()

    try:
        xml_files = extract_files([import_path], temp_dir)
        report = process_xml_files(xml_files, progress_dialog)

        official = load_official_keys()
        missing_keys = []
        if official:
            loaded_keys = set()
            for nota in report["notas"]:
                nNF = nota.get("nNF", "N/A")
                cNF = nota.get("cNF", "N/A")
                cnpj = nota.get("emitente", {}).get("cnpj", "N/A")
                loaded_keys.add((nNF, cNF, cnpj))
            missing_keys = list(official - loaded_keys)

        report["missing_keys"] = missing_keys

        logging.info(f"Arquivo '{file_path}' analisado.")
        logging.info(f"Total XML lidos: {len(xml_files)} | Notas válidas: {report['resumo']['total_notas']} | Erros: {len(report['errors'])} | Duplicadas: {len(report['duplicates'])}")
        if missing_keys:
            logging.info(f"Chaves ausentes: {len(missing_keys)}")

        return report
    finally:
        shutil.rmtree(temp_dir)

def extract_files(files: list, destination: str) -> list:
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
                    if filename.lower().endswith('.xml'):
                        full_path = os.path.join(root, filename)
                        shutil.copy(full_path, destination)

    extracted_files = [os.path.join(destination, f) for f in os.listdir(destination) if f.lower().endswith('.xml')]
    return extracted_files

def process_xml_files(xml_files: list, progress_dialog=None) -> dict:
    notas = []
    errors = []
    duplicates = []
    seen_keys = {}

    total_files = len(xml_files)

    for i, xml_file in enumerate(xml_files):
        try:
            nota_details = extract_note_details(xml_file)
            if nota_details:
                nNF = nota_details.get("nNF", "N/A")
                cNF = nota_details.get("cNF", "N/A")
                cnpj = nota_details.get("emitente", {}).get("cnpj", "N/A")
                key = (nNF, cNF, cnpj)
                if key in seen_keys:
                    duplicates.append(key)
                else:
                    seen_keys[key] = 1
                notas.append(nota_details)
        except ET.ParseError as e:
            errors.append(f"Erro parse '{xml_file}': {str(e)}")
        except Exception as e:
            errors.append(f"Erro process '{xml_file}': {str(e)}")

        if progress_dialog:
            progress_value = int((i + 1) / total_files * 100)
            progress_dialog.setValue(progress_value)
            if progress_dialog.wasCanceled():
                break

    resumo = {
        "total_notas": len(notas),
        "valor_total": sum(n.get("valor", 0.0) for n in notas)
    }

    return {
        "resumo": resumo,
        "notas": notas,
        "errors": errors,
        "duplicates": duplicates
    }

def extract_note_details(xml_file: str) -> dict:
    """
    Extrai os dados principais de uma nota a partir do XML.
    Também extrai o modelo com base no elemento <mod>:
      - Se <mod> for "55", define modelo como "NFE"
      - Se <mod> for "65", define modelo como "NFC-E"
    Se <mod> estiver ausente, usa o atributo Id de infNFe: se iniciar com "NFe", assume NFE; caso contrário, NFC-E.
    """
    detalhes = {
        "nome": os.path.basename(xml_file),
        "nNF": "N/A",
        "cNF": "N/A",
        "valor": 0.0,
        "status": "Desconhecido",
        "codigo_status": "",
        "autorizada": None,
        "emitida": None,
        "cancelada": False,
        "produtos": [],
        "emitente": {},
        "chNFe": None
    }

    namespace = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
    tree = ET.parse(xml_file)
    root = tree.getroot()

    infNFe = root.find(".//nfe:infNFe", namespace)
    if infNFe is not None:
        ch = infNFe.get("Id", "")
        if ch.startswith("NFe"):
            ch = ch[3:]
        detalhes["chNFe"] = ch

    protNFe = root.find(".//nfe:protNFe", namespace)

    nNF = root.find(".//nfe:ide/nfe:nNF", namespace)
    if nNF is not None:
        detalhes["nNF"] = nNF.text

    cNF = root.find(".//nfe:ide/nfe:cNF", namespace)
    if cNF is not None:
        detalhes["cNF"] = cNF.text

    # Extrai o modelo com base no elemento <mod>
    mod_elem = root.find(".//nfe:ide/nfe:mod", namespace)
    if mod_elem is not None and mod_elem.text:
        mod_val = mod_elem.text.strip()
        if mod_val == "55":
            detalhes["modelo"] = "NFE"
        elif mod_val == "65":
            detalhes["modelo"] = "NFC-E"
        else:
            detalhes["modelo"] = mod_val
    else:
        if infNFe is not None:
            id_val = infNFe.get("Id", "")
            if id_val.startswith("NFe"):
                detalhes["modelo"] = "NFE"
            else:
                detalhes["modelo"] = "NFC-E"
        else:
            detalhes["modelo"] = "NFC-E"

    vNF = root.find(".//nfe:vNF", namespace)
    if vNF is not None:
        try:
            detalhes["valor"] = float(vNF.text)
        except Exception as e:
            logging.error("Erro ao interpretar vNF em %s: %s", xml_file, e)
            detalhes["valor"] = 0.0

    cStat = root.find(".//nfe:cStat", namespace)
    if cStat is not None:
        detalhes["codigo_status"] = cStat.text
        if cStat.text in ["100", "150"]:
            if protNFe is not None:
                detalhes["status"] = "Autorizada"
            else:
                detalhes["status"] = "Sem Protocolo"
        elif cStat.text in ["101", "135", "151"]:
            detalhes["status"] = "Cancelada"
            detalhes["cancelada"] = True
        else:
            detalhes["status"] = "Desconhecido"

    dhRecbto = root.find(".//nfe:dhRecbto", namespace)
    if dhRecbto is not None:
        detalhes["autorizada"] = dhRecbto.text[:10]

    dhEmi = root.find(".//nfe:dhEmi", namespace)
    if dhEmi is not None:
        detalhes["emitida"] = dhEmi.text[:10]

    emitente = root.find(".//nfe:emit", namespace)
    if emitente is not None:
        nome = emitente.find("nfe:xNome", namespace)
        cnpj = emitente.find("nfe:CNPJ", namespace)
        ender = emitente.find("nfe:enderEmit", namespace)

        def get_text(tag):
            return tag.text if tag is not None else ""

        detalhes["emitente"] = {
            "nome": get_text(nome),
            "cnpj": get_text(cnpj),
            "endereco": ""
        }
        if ender is not None:
            xLgr = ender.find("nfe:xLgr", namespace)
            nro = ender.find("nfe:nro", namespace)
            bairro = ender.find("nfe:xBairro", namespace)
            xMun = ender.find("nfe:xMun", namespace)
            uf = ender.find("nfe:UF", namespace)
            detalhes["emitente"]["endereco"] = (
                f"{get_text(xLgr)}, {get_text(nro)}, {get_text(bairro)}, "
                f"{get_text(xMun)} - {get_text(uf)}"
            )

    for det in root.findall(".//nfe:det", namespace):
        prod = det.find("nfe:prod", namespace)
        if prod is not None:
            p = {
                "nome": "",
                "codigo": "",
                "cfop": "",
                "quantidade": 0.0,
                "unidade": "",
                "valor_unitario": 0.0,
                "valor_total": 0.0
            }
            xProd = prod.find("nfe:xProd", namespace)
            cProd = prod.find("nfe:cProd", namespace)
            cfop_tag = prod.find("nfe:CFOP", namespace)
            qCom = prod.find("nfe:qCom", namespace)
            uCom = prod.find("nfe:uCom", namespace)
            vUnCom = prod.find("nfe:vUnCom", namespace)
            vProd = prod.find("nfe:vProd", namespace)

            if xProd is not None:
                p["nome"] = xProd.text
            if cProd is not None:
                p["codigo"] = cProd.text
            if cfop_tag is not None:
                p["cfop"] = cfop_tag.text

            if qCom is not None:
                try:
                    p["quantidade"] = float(qCom.text)
                except Exception as e:
                    logging.error("Erro ao interpretar qCom em %s: %s", xml_file, e)
                    p["quantidade"] = 0.0
            if uCom is not None:
                p["unidade"] = uCom.text
            if vUnCom is not None:
                try:
                    p["valor_unitario"] = float(vUnCom.text)
                except Exception as e:
                    logging.error("Erro ao interpretar vUnCom em %s: %s", xml_file, e)
                    p["valor_unitario"] = 0.0
            if vProd is not None:
                try:
                    p["valor_total"] = float(vProd.text)
                except Exception as e:
                    logging.error("Erro ao interpretar vProd em %s: %s", xml_file, e)
                    p["valor_total"] = 0.0

            detalhes["produtos"].append(p)

    return detalhes
