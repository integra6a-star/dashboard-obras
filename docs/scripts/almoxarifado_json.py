# -*- coding: utf-8 -*-
"""Gera JSON do almoxarifado lendo diretamente as planilhas dos canteiros."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from datetime import datetime, timedelta
import json
import re
import unicodedata

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name.lower() in ("scripts", "pintura") else SCRIPT_DIR
SEARCH_DIRS = [ROOT, ROOT / "docs", SCRIPT_DIR]
OUT_DIR = ROOT / "docs" if (ROOT / "docs").exists() else ROOT

PLANILHAS = {
    "canteiro1": {"patterns": ["INTEGRA CANTEIRO 01*.xlsx", "CANTEIRO 01*.xlsx", "*CANTEIRO*01*.xlsx"], "saida": "almoxarifado_canteiro1.json"},
    "canteiro2": {"patterns": ["INTEGRA CANTEIRO 02*.xlsx", "CANTEIRO 02*.xlsx", "*CANTEIRO*02*.xlsx"], "saida": "almoxarifado_canteiro2.json"},
    "iti15": {"patterns": ["INTEGRA CANTEIRO ITI-15*.xlsx", "INTEGRA CANTEIRO ITI15*.xlsx", "CANTEIRO ITI-15*.xlsx", "CANTEIRO ITI15*.xlsx", "*ITI-15*.xlsx", "*ITI15*.xlsx"], "saida": "almoxarifado_iti15.json"},
}

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn")


def limpar_descricao(txt):
    txt = re.sub(r"\s+", " ", str(txt or "").strip())
    for errado, certo in {"ADOELA": "ADUELA", "Adoela": "Aduela", "adoela": "aduela", "CURRUGADO": "CORRUGADO", "Surperior": "Superior", "SURPERIOR": "SUPERIOR"}.items():
        txt = txt.replace(errado, certo)
    return txt


def numero(valor):
    if valor is None or valor == "":
        return 0
    if isinstance(valor, (int, float)):
        return int(valor) if float(valor).is_integer() else float(valor)
    txt = str(valor).strip().replace("R$", "").replace(" ", "")
    if txt in ["-", ".", ","]:
        return 0
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    try:
        n = float(txt)
        return int(n) if n.is_integer() else n
    except Exception:
        return 0


def data_txt(valor):
    if valor is None or valor == "":
        return ""
    try:
        n = float(valor)
        if 20000 < n < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=n)).strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(valor).strip()


def normalizar_tipo(tipo):
    t = sem_acento(tipo).strip().lower()
    if "said" in t or t.startswith("sai") or t == "s":
        return "Saída"
    if "entr" in t or t == "e":
        return "Entrada"
    return str(tipo or "").strip()


def encontrar_planilha(info):
    candidatos = []
    for pasta in SEARCH_DIRS:
        if pasta.exists():
            for pattern in info["patterns"]:
                candidatos.extend(pasta.glob(pattern))
    candidatos = [p for p in candidatos if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall(NS + "si"):
        out.append("".join((t.text or "") for t in si.iter(NS + "t")))
    return out


def mapa_abas(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    abas = {}
    for sheet in wb.find(NS + "sheets"):
        rid = sheet.attrib[NS_REL + "id"]
        alvo = relmap[rid]
        caminho = alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo
        abas[sheet.attrib["name"]] = caminho
    return abas


def col_idx(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return None
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_value(c, ss):
    tipo = c.attrib.get("t")
    if tipo == "inlineStr":
        return "".join((t.text or "") for t in c.iter(NS + "t"))
    v = c.find(NS + "v")
    txt = v.text if v is not None else ""
    if tipo == "s":
        try:
            return ss[int(txt)]
        except Exception:
            return ""
    return txt


def ler_aba(caminho, nome_aba):
    with ZipFile(caminho) as z:
        abas = mapa_abas(z)
        alvo = None
        for nome, path in abas.items():
            if sem_acento(nome).lower() == sem_acento(nome_aba).lower():
                alvo = path
                break
        if not alvo:
            return []
        ss = shared_strings(z)
        linhas = []
        for _, row in ET.iterparse(z.open(alvo), events=("end",)):
            if row.tag != NS + "row":
                continue
            vals = {}
            max_col = -1
            for c in row.findall(NS + "c"):
                idx = col_idx(c.attrib.get("r", ""))
                if idx is None:
                    continue
                vals[idx] = cell_value(c, ss)
                max_col = max(max_col, idx)
            linhas.append([vals.get(i, "") for i in range(max_col + 1)])
            row.clear()
        return linhas


def processar_planilha(caminho):
    produtos = {}
    for row in ler_aba(caminho, "Produtos")[3:]:
        codigo = row[1] if len(row) > 1 else ""
        descricao = row[2] if len(row) > 2 else ""
        if not codigo and not descricao:
            continue
        codigo = str(codigo or "").strip()
        descricao = limpar_descricao(descricao)
        if not codigo or not descricao:
            continue
        produtos[codigo] = {
            "codigo": codigo,
            "descricao": descricao,
            "estoque_minimo": numero(row[3] if len(row) > 3 else 0),
            "entradas": 0,
            "saidas": 0,
            "quantidade": numero(row[4] if len(row) > 4 else 0),
            "situacao_planilha": str(row[5] if len(row) > 5 else "").strip(),
        }

    movimentos = []
    for row in ler_aba(caminho, "Movimentos")[3:]:
        data = row[1] if len(row) > 1 else ""
        tipo = row[2] if len(row) > 2 else ""
        codigo = str(row[3] if len(row) > 3 else "").strip()
        descricao = row[4] if len(row) > 4 else ""
        qtd = numero(row[5] if len(row) > 5 else 0)
        requisitante = row[6] if len(row) > 6 else ""
        obs = row[7] if len(row) > 7 else ""
        if not codigo or qtd == 0:
            continue
        tipo_norm = normalizar_tipo(tipo)
        if codigo not in produtos:
            produtos[codigo] = {"codigo": codigo, "descricao": limpar_descricao(descricao), "estoque_minimo": 0, "entradas": 0, "saidas": 0, "quantidade": 0, "situacao_planilha": ""}
        if tipo_norm == "Entrada":
            produtos[codigo]["entradas"] += qtd
        elif tipo_norm == "Saída":
            produtos[codigo]["saidas"] += qtd
        movimentos.append({"data": data_txt(data), "tipo": tipo_norm, "codigo": codigo, "descricao": limpar_descricao(descricao) or produtos[codigo]["descricao"], "quantidade": qtd, "requisitante": str(requisitante or "").strip(), "obs": str(obs or "").strip()})

    return {"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "origem": caminho.name, "produtos": list(produtos.values()), "movimentos": movimentos}


def salvar_json(nome, dados):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    destino = OUT_DIR / nome
    destino.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    return destino


def main():
    print("Gerando JSONs do Almoxarifado...")
    saidas = {}
    for chave, info in PLANILHAS.items():
        planilha = encontrar_planilha(info)
        if not planilha:
            dados = {"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "origem": "", "produtos": [], "movimentos": []}
            destino = salvar_json(info["saida"], dados)
            print(f"AVISO: {chave}: planilha não encontrada -> {destino.name}")
        else:
            dados = processar_planilha(planilha)
            destino = salvar_json(info["saida"], dados)
            print(f"OK: {chave}: {planilha.name} -> {destino.name} | produtos: {len(dados['produtos'])} | movimentos: {len(dados['movimentos'])}")
        saidas[chave] = dados
    salvar_json("almoxarifado.json", saidas.get("canteiro1", {"produtos": [], "movimentos": []}))
    print("Concluído.")

if __name__ == "__main__":
    main()
