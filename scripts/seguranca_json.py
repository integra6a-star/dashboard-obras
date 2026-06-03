from __future__ import annotations

import json
import re
import unicodedata
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


def normalize_key(value: str) -> str:
    value = unicodedata.normalize("NFD", value or "")
    value = "".join(char for char in value if unicodedata.category(char) != "Mn")
    return re.sub(r"\s+", " ", value).strip().lower()


def get_sheet_path(sheets: dict[str, str], *names: str) -> str | None:
    normalized = {normalize_key(name): path for name, path in sheets.items()}
    for name in names:
        path = normalized.get(normalize_key(name))
        if path:
            return path
    return None


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
    value = value.replace("•", ";").replace("â€¢", ";")
    parts = [clean(part, "") for part in re.split(r";+|\n+", value)]
    return [part for part in parts if part]


def is_header_or_empty(row: list[str]) -> bool:
    first = clean(row[0] if row else "").upper()
    return not first or first in {"DATA", "NOME", "COLABORADOR", "TOTAL"}


def build_classification_lookup(rows: list[list[str]]) -> dict[str, dict[str, str]]:
    group_columns = [(9, 10), (12, 13), (15, 16), (18, 19), (21, 22)]
    lookup: dict[str, dict[str, str]] = {}
    for code_col, desc_col in group_columns:
        group = ""
        group_letter = ""
        for row in rows:
            code_cell = clean(row[code_col] if code_col < len(row) else "")
            desc_cell = clean(row[desc_col] if desc_col < len(row) else "")
            group_match = re.match(r"^([A-Z])\.\s+(.+)$", code_cell)
            item_match = re.match(r"^([A-Z]\d+)\.?", code_cell)
            if group_match and not item_match:
                group_letter = group_match.group(1)
                group = code_cell
                continue
            if item_match:
                code = item_match.group(1).upper()
                lookup[code] = {
                    "codigo": code,
                    "grupo_codigo": group_letter or code[0],
                    "grupo": group or f"{code[0]}. Não classificado",
                    "descricao": desc_cell,
                    "texto": clean(f"{code}. {desc_cell}"),
                }
    return lookup


def parse_classificacoes(value: str, lookup: dict[str, dict[str, str]]) -> list[dict[str, str]]:
    classificacoes = []
    value = clean(value)
    matches = list(re.finditer(r"\b([ABCDF]\d+)\.?", value, flags=re.IGNORECASE))
    for index, match in enumerate(matches):
        code = match.group(1).upper()
        next_start = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        descricao = value[match.end() : next_start]
        descricao = re.sub(r"^[\s.;:/\\-]+", "", descricao)
        descricao = re.sub(r"[\s.;:/\\-]+$", "", descricao)
        descricao = clean(descricao)
        base = lookup.get(code, {})
        final_descricao = base.get("descricao", "") or descricao
        classificacoes.append(
            {
                "codigo": code,
                "grupo_codigo": base.get("grupo_codigo", code[0]),
                "grupo": base.get("grupo", f"{code[0]}. Nao classificado"),
                "descricao": final_descricao,
                "texto": clean(f"{code}. {final_descricao}"),
            }
        )
    return classificacoes


def join_unique(values: list[str], fallback: str) -> str:
    unique = []
    for value in values:
        value = clean(value)
        if value and value not in unique:
            unique.append(value)
    return "; ".join(unique) if unique else fallback


def parse_inspecoes(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]):
    all_rows = list(iter_rows(zf, sheet_path, shared))
    header = all_rows[0] if all_rows else []
    rows = all_rows[1:]
    lookup = {clean(name).upper(): idx for idx, name in enumerate(header)}
    classificacao_lookup = build_classification_lookup(rows)

    def get(row: list[str], name: str, fallback_idx: int | None = None) -> str:
        idx = lookup.get(name)
        if idx is None and fallback_idx is not None:
            idx = fallback_idx
        return row[idx] if idx is not None and idx < len(row) else ""

    inspecoes = []
    desvios = []
    for row in rows:
        if is_header_or_empty(row):
            continue
        data = excel_date(get(row, "DATA", 0))
        local = clean(get(row, "LOCAL", 1), "Sem localização")
        lider = clean(get(row, "ENCARREGADO OU LIDER DE EQUIPE", 2), "Não informado")
        descricao = clean(get(row, "DESCRIÇÃO", 4))
        classificacao_original = clean(get(row, "CLASSIFICAÇÃO", 5), "Não classificado")
        classificacoes = parse_classificacoes(classificacao_original, classificacao_lookup)
        categoria = join_unique([item["grupo"] for item in classificacoes], "Não classificado")
        categoria_detalhe = join_unique([item["texto"] for item in classificacoes], classificacao_original)
        qtd_desvios = integer(get(row, "QUANTIDADE DESVIOS", 3))
        qtd_inspecoes = integer(get(row, "QUANTIDADE INSPEÇÕES", 6)) or 1
        tst = clean(get(row, "TÉCNICO DE SEGURANÇA  APLICADOR", 7), "Não informado")

        if not data and not descricao and local == "Sem localização":
            continue

        inspecoes.append(
            {
                "data": data,
                "local": local,
                "lider": lider,
                "descricao": descricao,
                "categoria": categoria,
                "categoria_detalhe": categoria_detalhe,
                "classificacao_original": classificacao_original,
                "classificacoes": classificacoes,
                "tst": tst,
                "quantidade_desvios": qtd_desvios,
                "quantidade_inspecoes": qtd_inspecoes,
            }
        )

        if classificacoes:
            for index, classificacao in enumerate(classificacoes, start=1):
                desvios.append(
                    {
                        "data": data,
                        "local": local,
                        "lider": lider,
                        "tst": tst,
                        "categoria": classificacao["texto"],
                        "categoria_grupo": classificacao["grupo"],
                        "categoria_detalhe": classificacao["texto"],
                        "classificacao_codigo": classificacao["codigo"],
                        "classificacao_grupo": classificacao["grupo"],
                        "classificacao_descricao": classificacao["descricao"],
                        "classificacao_original": classificacao_original,
                        "desvio": descricao,
                        "descricao": descricao,
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
        inspecoes_path = get_sheet_path(sheets, "Inspeções", "Inspecoes")
        dds_path = get_sheet_path(sheets, "DDS")
        treinamentos_path = get_sheet_path(sheets, "Treinamentos admissionais")
        listas_path = get_sheet_path(sheets, "Listas")
        if not inspecoes_path:
            raise KeyError("Aba Inspeções não encontrada.")
        inspecoes, desvios = parse_inspecoes(zf, inspecoes_path, shared)
        dds = parse_dds(zf, dds_path, shared) if dds_path else []
        treinamentos = count_treinamentos(zf, treinamentos_path, shared) if treinamentos_path else 0
        listas = parse_listas(zf, listas_path, shared) if listas_path else {}

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
