import json
from pathlib import Path
from datetime import datetime, date
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

# =========================
# CAMINHOS (ROBUSTO)
# =========================
SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent  # pasta do projeto (dashboard-obras)

EXCEL_NAME = "BASE_DASH_EXTENSAO_POWERBI.xlsx"

# tenta achar o Excel em locais comuns
CANDIDATOS_EXCEL = [
    ROOT_DIR / "docs" / EXCEL_NAME,   # ✅ padrão: excel dentro de docs
    ROOT_DIR / EXCEL_NAME,            # excel na raiz
    SCRIPT_DIR / EXCEL_NAME,          # excel dentro de scripts
]

ARQ_EXCEL = next((p for p in CANDIDATOS_EXCEL if p.exists()), None)
SAIDA_JSON = (ROOT_DIR / "docs" / "dados.json").resolve()

# =========================
# MAPEAMENTO DE COLUNAS
# =========================
EXPECTED = {
    "Data": ["data"],
    "Obra": ["obra"],
    "Bloco": ["bloco"],
    "Tipo": ["tipo_extensao", "tipo extensao", "tipo", "tipoextensao"],
    "Planejado_m": ["extensao_planejada_m", "extensao planejada (m)", "extensao planejada", "planejado_m"],
    "Executado_m": ["extensao_executada_m", "extensao executada (m)", "extensao executada", "executado_m"],
    "PV": ["pv", "pvs", "qtd pv", "quantidade pv", "qtdpv"],
    "Profundidade_m": ["profundidade_pv_m", "profundidade pv (m)", "profundidade", "profundidade_m"],
    "Economias_Previstas": ["economias prevista", "economias previstas", "econ prev", "economias prevista(s)"],
    "Economias_Recebidas": ["economias recebidas", "economias recebida", "econ receb", "economias recebida(s)"],
}

def norm(s):
    return "".join(str(s).strip().lower().split())

def to_float(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def to_iso_date(v):
    """
    Converte Data para 'YYYY-MM-DD' (OBRIGATÓRIO para Curva S).
    Aceita:
    - datetime/date do Excel
    - texto 'dd/mm/aaaa' ou 'yyyy-mm-dd'
    """
    if v is None or v == "":
        return ""

    if isinstance(v, datetime):
        return v.date().isoformat()

    if isinstance(v, date):
        return v.isoformat()

    s = str(v).strip()

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except:
            pass

    return ""  # se não conseguir converter

def main():
    if ARQ_EXCEL is None:
        raise FileNotFoundError(
            f"❌ Não encontrei o Excel '{EXCEL_NAME}'.\n"
            f"➡️ Coloque ele em: {ROOT_DIR / 'docs'} (recomendado) ou na raiz do projeto."
        )

    wb = load_workbook(ARQ_EXCEL, data_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = [cell.value for cell in ws[1]]
    if not headers or all(h is None for h in headers):
        raise ValueError("❌ Primeira linha da planilha está vazia (sem cabeçalhos).")

    headers_norm = [norm(h) if h is not None else "" for h in headers]

    col_index = {}
    for out_name, candidates in EXPECTED.items():
        found = None
        for cand in candidates:
            cand_n = norm(cand)
            for i, hn in enumerate(headers_norm):
                if hn == cand_n:
                    found = i
                    break
            if found is not None:
                break

        if found is None:
            raise ValueError(
                f"❌ Não encontrei a coluna para '{out_name}'.\n"
                f"Cabeçalhos encontrados: {headers}"
            )

        col_index[out_name] = found

    registros = []
    datas_vazias = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None for v in row):
            continue

        d_iso = to_iso_date(row[col_index["Data"]])
        if not d_iso:
            datas_vazias += 1

        registros.append({
            "Data": d_iso,  # ✅ agora a Curva S tem data!
            "Obra": str(row[col_index["Obra"]] or "").strip(),
            "Bloco": str(row[col_index["Bloco"]] or "").strip(),
            "Tipo": str(row[col_index["Tipo"]] or "").strip(),
            "Planejado_m": to_float(row[col_index["Planejado_m"]]),
            "Executado_m": to_float(row[col_index["Executado_m"]]),
            "PV": to_float(row[col_index["PV"]]),
            "Profundidade_m": to_float(row[col_index["Profundidade_m"]]),
            "Economias_Previstas": to_float(row[col_index["Economias_Previstas"]]),
            "Economias_Recebidas": to_float(row[col_index["Economias_Recebidas"]]),
        })

    payload = {
        "atualizado_em": datetime.now(ZoneInfo("America/Sao_Paulo")).isoformat(),
        "registros": registros
    }

    SAIDA_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(SAIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"✅ OK: {SAIDA_JSON} gerado com {len(registros)} linhas.")
    print(f"📌 Excel lido de: {ARQ_EXCEL}")
    if datas_vazias > 0:
        print(f"⚠️ Atenção: {datas_vazias} linhas ficaram com Data vazia (Curva S ignora essas linhas).")

if __name__ == "__main__":
    main()
