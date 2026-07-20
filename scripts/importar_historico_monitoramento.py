import json
import re
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pdfplumber


ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
OUT_ROOT = ROOT / "monitoramento_historico.json"
OUT_DOCS = DOCS / "monitoramento_historico.json"

ZIPS = [
    Path(r"C:\Users\micro\Downloads\OneDrive_2026-07-20.zip"),
    Path(r"C:\Users\micro\Downloads\OneDrive_2026-07-20 (1).zip"),
    Path(r"C:\Users\micro\Downloads\OneDrive_2026-07-20 (2).zip"),
]

MESES = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def parse_float(value):
    if value is None:
        return None
    text = str(value).replace("+", "").replace(",", ".").strip()
    try:
        return float(text)
    except ValueError:
        return None


def parse_date_from_name(name):
    match = re.search(r"_(\d{2})_([a-z]{3})\.pdf$", name, re.I)
    if not match:
        return None
    day = int(match.group(1))
    month = MESES.get(match.group(2).lower())
    if not month:
        return None
    return f"2026-{month:02d}-{day:02d}"


def normalize_poco(raw):
    text = re.sub(r"\s+", " ", raw or "").strip().upper()
    text = text.replace("_", " ")
    if "POSTE" in text:
        match = re.search(r"PV\s*-?\s*0*(\d+)", text)
        return f"POSTE PV-{int(match.group(1)):02d}" if match else "POSTE PV-06"
    match = re.search(r"PV\s*-?\s*0*(\d+)", text)
    if match:
        return f"PV-{int(match.group(1)):02d}"
    return text


def parse_poco_from_text_or_name(text, name):
    header = re.search(r"\|\s*([^|]+?)\s*\|\s*ULTIMOS", text, re.I)
    if header:
        return normalize_poco(header.group(1))
    name_match = re.search(r"(POSTE[_\s]*PV[_\s-]*0*\d+|PV[_\s-]*0*\d+)", name, re.I)
    return normalize_poco(name_match.group(1) if name_match else name)


def report_type(name, text):
    source = f"{name}\n{text}".upper()
    if "VARIACAO_MM" in source or "VARIAÇÃO EM MM" in source:
        return "variacao_mm"
    if "ABSOLUTO" in source or "VARIAÇÃO ABSOLUTA" in source:
        return "absoluto"
    if "DESVIO" in source:
        return "desvio"
    return "outro"


def parse_lines(text):
    parts = re.split(r"(?=LINHA\s+)", text)
    lines = []
    for part in parts:
        head = re.search(r"LINHA\s+([A-Z0-9]+)(?:\s+-\s+REFER|\s|$)", part, re.I)
        if not head:
            continue
        line = head.group(1).upper()
        if not (re.fullmatch(r"L\d+", line) or line == "POSTE"):
            continue
        summary = re.search(
            r"MEDIÇÕES:\s*(\d+)(?:\s+M[ÉE]DIA:\s*([+\-]?\d+(?:[,.]\d+)?)\s*MM)?\s+MENOR:\s*([+\-]?\d+(?:[,.]\d+)?)\s*MM?\s+MAIOR:\s*([+\-]?\d+(?:[,.]\d+)?)",
            part,
            re.I,
        )
        if not summary:
            summary = re.search(
                r"MEDIÇÕES:\s*(\d+)\s+MENOR:\s*([+\-]?\d+(?:[,.]\d+)?)\s+MAIOR:\s*([+\-]?\d+(?:[,.]\d+)?)",
                part,
                re.I,
            )
            if not summary:
                continue
            medicoes = int(summary.group(1))
            media = None
            menor = parse_float(summary.group(2))
            maior = parse_float(summary.group(3))
        else:
            medicoes = int(summary.group(1))
            media = parse_float(summary.group(2))
            menor = parse_float(summary.group(3))
            maior = parse_float(summary.group(4))

        trechos = []
        trecho_rx = re.compile(
            rf"({line}\s+-\s+A\d\s+PARA\s+A\d)\s+AMPLITUDE:\s*([+\-]?\d+(?:[,.]\d+)?)\s*MM\s+DESLOCAMENTO:\s*([+\-]?\d+(?:[,.]\d+)?)\s*MM",
            re.I,
        )
        for trecho, amp, desl in trecho_rx.findall(part):
            trechos.append({
                "nome": re.sub(r"\s+", " ", trecho).upper(),
                "amplitude_mm": parse_float(amp),
                "deslocamento_mm": parse_float(desl),
            })

        lines.append({
            "nome": line,
            "medicoes": medicoes,
            "media_mm": media,
            "menor_mm": menor,
            "maior_mm": maior,
            "amplitude_mm": max(abs(v) for v in (menor or 0, maior or 0)),
            "trechos": trechos,
        })
    return lines


def extract_pdf_text(zip_path, entry_name):
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / Path(entry_name).name
        with zipfile.ZipFile(zip_path) as zf:
            pdf_path.write_bytes(zf.read(entry_name))
        with pdfplumber.open(str(pdf_path)) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)


def main():
    registros = []
    erros = []
    for zip_path in ZIPS:
        if not zip_path.exists():
            erros.append(f"ZIP nao encontrado: {zip_path}")
            continue
        with zipfile.ZipFile(zip_path) as zf:
            entries = [e for e in zf.namelist() if e.lower().endswith(".pdf")]
        for entry in entries:
            try:
                data = parse_date_from_name(entry)
                if not data:
                    continue
                text = extract_pdf_text(zip_path, entry)
                tipo = report_type(entry, text)
                if tipo not in {"variacao_mm", "desvio", "absoluto"}:
                    continue
                linhas = parse_lines(text)
                if not linhas:
                    continue
                registros.append({
                    "data": data,
                    "poco": parse_poco_from_text_or_name(text, entry),
                    "tipo": tipo,
                    "arquivo": entry,
                    "linhas": linhas,
                })
            except Exception as exc:
                erros.append(f"{entry}: {exc}")

    por_poco = {}
    for registro in registros:
        poco = por_poco.setdefault(registro["poco"], {"poco": registro["poco"], "relatorios": [], "linhas": {}})
        poco["relatorios"].append(registro)
        for linha in registro["linhas"]:
            item = poco["linhas"].setdefault(linha["nome"], {
                "nome": linha["nome"],
                "serie": [],
                "trechos": [],
                "medicoes": 0,
            })
            item["serie"].append({
                "data": registro["data"],
                "tipo": registro["tipo"],
                "media_mm": linha["media_mm"],
                "menor_mm": linha["menor_mm"],
                "maior_mm": linha["maior_mm"],
                "amplitude_mm": linha["amplitude_mm"],
                "medicoes": linha["medicoes"],
            })
            if registro["tipo"] == "variacao_mm" and linha["trechos"]:
                item["trechos"] = linha["trechos"]
            item["medicoes"] = max(item["medicoes"], linha["medicoes"] or 0)

    pocos = []
    for poco in por_poco.values():
        linhas = []
        for linha in poco["linhas"].values():
            linha["serie"].sort(key=lambda row: (row["data"], row["tipo"]))
            linhas.append(linha)
        linhas.sort(key=lambda row: row["nome"])
        relatorios = sorted(poco["relatorios"], key=lambda row: (row["data"], row["tipo"], row["arquivo"]))
        pocos.append({
            "poco": poco["poco"],
            "linhas": linhas,
            "total_relatorios": len(relatorios),
            "ultima_data": relatorios[-1]["data"] if relatorios else None,
        })
    pocos.sort(key=lambda row: row["poco"])

    payload = {
        "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "fonte": "PDFs historicos de monitoramento topografico",
        "total_relatorios": len(registros),
        "total_pocos": len(pocos),
        "pocos": pocos,
        "erros": erros[:50],
    }
    OUT_ROOT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    OUT_DOCS.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Historico salvo: {len(registros)} relatorio(s), {len(pocos)} poco(s).")
    if erros:
        print(f"Avisos: {len(erros)} arquivo(s) nao importado(s).")


if __name__ == "__main__":
    main()
