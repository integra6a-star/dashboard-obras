from __future__ import annotations

import json
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
NAMESPACE = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}

SOURCE_CANDIDATES = [
    ROOT / "Planilha de controle de seguranca do trabalho.xlsx",
    ROOT / "Planilha de controle de segurança do trabalho (2).xlsx",
    Path(r"C:\Users\micro\Downloads\Planilha de controle de segurança do trabalho (2).xlsx"),
]


def text_of(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    return [text_of(si) for si in root.findall("main:si", NAMESPACE)]


def workbook_sheets(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkg:Relationship", NAMESPACE)
    }
    sheets: dict[str, str] = {}
    for sheet in workbook.findall("main:sheets/main:sheet", NAMESPACE):
        rid = sheet.attrib.get(f"{{{NAMESPACE['rel']}}}id")
        target = rel_targets.get(rid or "")
        if target:
            sheets[sheet.attrib["name"]] = "xl/" + target.lstrip("/")
    return sheets


def column_index(cell_ref: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
    idx = 0
    for letter in letters:
        idx = idx * 26 + (ord(letter) - 64)
    return idx - 1


def cell_value(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", NAMESPACE)
    if cell_type == "inlineStr":
        return text_of(cell.find("main:is", NAMESPACE))
    raw = text_of(value)
    if cell_type == "s" and raw:
        try:
            return shared[int(float(raw))]
        except (IndexError, ValueError):
            return raw
    return raw


def iter_rows(zf: zipfile.ZipFile, path: str, shared: list[str]):
    with zf.open(path) as fh:
        for event, row in ET.iterparse(fh, events=("end",)):
            if not row.tag.endswith("row"):
                continue
            values: dict[int, str] = {}
            for cell in row.findall("main:c", NAMESPACE):
                ref = cell.attrib.get("r", "")
                if not ref:
                    continue
                values[column_index(ref)] = cell_value(cell, shared)
            if values:
                max_idx = max(values)
                yield [values.get(i, "").strip() for i in range(max_idx + 1)]
            row.clear()


def excel_date(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    try:
        serial = float(value.replace(",", "."))
        return (datetime(1899, 12, 30) + timedelta(days=serial)).date().isoformat()
    except ValueError:
        pass
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue
    return value


def clean(value: str, fallback: str = "") -> str:
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    return value or fallback


def integer(value: str) -> int:
    try:
        return int(float(str(value).replace(",", ".")))
    except ValueError:
        return 0


def split_parts(value: str) -> list[str]:
    value = clean(value)
    if not value:
        return []
    value = value.replace("•", ";")
    parts = [clean(part, "") for part in re.split(r";+|\n+", value)]
    return [part for part in parts if part]


def is_header_or_empty(row: list[str]) -> bool:
    first = clean(row[0] if row else "").upper()
    return not first or first in {"DATA", "NOME", "COLABORADOR", "TOTAL"}


def parse_inspecoes(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]):
    rows = iter_rows(zf, sheet_path, shared)
    header = next(rows, [])
    lookup = {clean(name).upper(): idx for idx, name in enumerate(header)}

    def get(row: list[str], name: str) -> str:
        idx = lookup.get(name)
        return row[idx] if idx is not None and idx < len(row) else ""

    inspecoes = []
    desvios = []
    for row in rows:
        if is_header_or_empty(row):
            continue
        data = excel_date(get(row, "DATA"))
        local = clean(get(row, "LOCAL"), "Sem localização")
        lider = clean(get(row, "ENCARREGADO OU LIDER DE EQUIPE"), "Não informado")
        descricao = clean(get(row, "DESCRIÇÃO"))
        categoria = clean(get(row, "CLASSIFICAÇÃO"), "Não classificado")
        qtd_desvios = integer(get(row, "QUANTIDADE DESVIOS"))
        qtd_inspecoes = integer(get(row, "QUANTIDADE INSPEÇÕES")) or 1
        tst = clean(get(row, "TÉCNICO DE SEGURANÇA  APLICADOR"), "Não informado")

        if not data and not descricao and local == "Sem localização":
            continue

        inspecoes.append(
            {
                "data": data,
                "local": local,
                "lider": lider,
                "descricao": descricao,
                "categoria": categoria,
                "tst": tst,
                "quantidade_desvios": qtd_desvios,
                "quantidade_inspecoes": qtd_inspecoes,
            }
        )

        descricao_parts = split_parts(descricao)
        categoria_parts = split_parts(categoria)
        if descricao_parts or qtd_desvios > 0:
            if not descricao_parts:
                descricao_parts = [descricao or "Desvio sem descrição"]
            for index, desvio in enumerate(descricao_parts, start=1):
                desvios.append(
                    {
                        "data": data,
                        "local": local,
                        "lider": lider,
                        "tst": tst,
                        "categoria": categoria_parts[index - 1] if index - 1 < len(categoria_parts) else categoria,
                        "desvio": desvio,
                        "descricao": desvio,
                        "quantidade_desvios": qtd_desvios,
                        "ordem_desvio": index,
                    }
                )

    return inspecoes, desvios


def parse_dds(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]):
    rows = iter_rows(zf, sheet_path, shared)
    header = next(rows, [])
    lookup = {clean(name).upper(): idx for idx, name in enumerate(header)}

    def get(row: list[str], name: str) -> str:
        idx = lookup.get(name)
        return row[idx] if idx is not None and idx < len(row) else ""

    dds = []
    for row in rows:
        if is_header_or_empty(row):
            continue
        data = excel_date(get(row, "DATA"))
        tema = clean(get(row, "TEMA"))
        responsavel = clean(get(row, "RESPONSÁVEL"), "Responsável não identificado")
        participantes = integer(get(row, "QUANTIDADE DE PARTICIPANTES"))
        if data and tema:
            dds.append(
                {
                    "data": data,
                    "tema": tema,
                    "responsavel": responsavel,
                    "participantes": participantes,
                }
            )
    return dds


def count_treinamentos(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> int:
    total = 0
    for row in iter_rows(zf, sheet_path, shared):
        nome = clean(row[0] if len(row) > 0 else "")
        cpf = clean(row[1] if len(row) > 1 else "")
        funcao = clean(row[2] if len(row) > 2 else "")
        if not nome or nome.upper() in {"NOME", "COLABORADOR"}:
            continue
        if len(re.sub(r"\D", "", cpf)) >= 8 and funcao:
            total += 1
    return total


def parse_listas(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]):
    rows = iter_rows(zf, sheet_path, shared)
    header = next(rows, [])
    listas = {clean(name): [] for name in header if clean(name)}
    for row in rows:
        for idx, name in enumerate(header):
            key = clean(name)
            if key and idx < len(row) and clean(row[idx]):
                listas[key].append(clean(row[idx]))
    return listas


def build_payload(source: Path):
    with zipfile.ZipFile(source) as zf:
        shared = load_shared_strings(zf)
        sheets = workbook_sheets(zf)
        inspecoes, desvios = parse_inspecoes(zf, sheets["Inspeções"], shared)
        dds = parse_dds(zf, sheets["DDS"], shared)
        treinamentos = count_treinamentos(zf, sheets["Treinamentos admissionais"], shared)
        listas = parse_listas(zf, sheets["Listas"], shared)

    return {
        "metadata": {
            "fonte": source.name,
            "gerado_em": datetime.now().isoformat(timespec="seconds"),
            "estrutura": "Planilha de controle de segurança do trabalho",
        },
        "inspecoes": inspecoes,
        "desvios": desvios,
        "dds": dds,
        "treinamentosAdmissao": treinamentos,
        "listas": listas,
    }


def main() -> None:
    source = next((path for path in SOURCE_CANDIDATES if path.exists()), None)
    if not source:
        raise FileNotFoundError("Planilha de segurança não encontrada.")
    payload = build_payload(source)
    for target in [ROOT / "seguranca_data.json", ROOT / "docs" / "seguranca_data.json"]:
        target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        "Segurança atualizada:",
        f"inspeções={len(payload['inspecoes'])}",
        f"desvios={len(payload['desvios'])}",
        f"dds={len(payload['dds'])}",
        f"treinamentos={payload['treinamentosAdmissao']}",
    )


if __name__ == "__main__":
    main()
