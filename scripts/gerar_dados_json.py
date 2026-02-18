import json
import pandas as pd
from pathlib import Path

# Nome do arquivo Excel (já está no seu repo)
ARQ_EXCEL = Path("BASE_DASH_EXTENSAO_POWERBI.xlsx")

# Onde o JSON será criado
SAIDA_JSON = Path("docs/dados.json")

# Mapeamento de possíveis nomes de colunas (para evitar erro)
COLS_MAP = {
    "Obra": ["Obra", "OBRA"],
    "Bloco": ["Bloco", "BLOCO"],
    "Tipo_Extensao": ["Tipo_Extensao", "Tipo Extensao", "Tipo", "TIPO_EXTENSAO"],
    "Extensao_Planejada_m": ["Extensao_Planejada_m", "Extensão Planejada (m)", "Extensao Planejada", "Planejado_m"],
    "Extensao_Executada_m": ["Extensao_Executada_m", "Extensão Executada (m)", "Extensao Executada", "Executado_m"],
    "PV": ["PV", "PVs", "Qtd PV", "Quantidade PV"],
    "Profundidade_PV_m": ["Profundidade_PV_m", "Profundidade PV (m)", "Profundidade", "Profundidade_m"],
    "Economias prevista": ["Economias prevista", "Economias Prevista", "Economias Previstas", "Econ Prev"],
    "Economias Recebidas": ["Economias Recebidas", "Economias Recebida", "Econ Receb"]
}

def pick_col(df, wanted):
    for cand in COLS_MAP[wanted]:
        if cand in df.columns:
            return cand
    return None

def main():
    if not ARQ_EXCEL.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQ_EXCEL}")

    # Lê a primeira aba do Excel
    xls = pd.ExcelFile(ARQ_EXCEL)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(ARQ_EXCEL, sheet_name=sheet)

    # Identifica as colunas corretas
    real = {}
    for key in COLS_MAP.keys():
        col = pick_col(df, key)
        if col is None:
            raise ValueError(
                f"Coluna obrigatória não encontrada: {key}. "
                f"Colunas atuais: {list(df.columns)}"
            )
        real[key] = col

    # Seleciona e padroniza colunas
    out = df[
        [
            real["Obra"],
            real["Bloco"],
            real["Tipo_Extensao"],
            real["Extensao_Planejada_m"],
            real["Extensao_Executada_m"],
            real["PV"],
            real["Profundidade_PV_m"],
            real["Economias prevista"],
            real["Economias Recebidas"],
        ]
    ].copy()

    out.columns = [
        "Obra",
        "Bloco",
        "Tipo",
        "Planejado_m",
        "Executado_m",
        "PV",
        "Profundidade_m",
        "Economias_Previstas",
        "Economias_Recebidas",
    ]

    # Converte colunas numéricas
    num_cols = [
        "Planejado_m",
        "Executado_m",
        "PV",
        "Profundidade_m",
        "Economias_Previstas",
        "Economias_Recebidas",
    ]
    for c in num_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)

    # Monta o JSON final
    payload = {
        "atualizado_em": pd.Timestamp.now(tz="America/Sao_Paulo").isoformat(),
        "registros": out.to_dict(orient="records"),
    }

    SAIDA_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(SAIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"OK: gerado {SAIDA_JSON} com {len(out)} linhas (aba: {sheet}).")

if __name__ == "__main__":
    main()
