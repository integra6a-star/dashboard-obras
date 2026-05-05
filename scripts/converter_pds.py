import json
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DOCS_DIR = BASE_DIR.parent / "docs"
ARQUIVO_XLSX = DOCS_DIR / "pds_word_completo_preenchido(2).xlsx"
ARQUIVO_JSON = DOCS_DIR / "pds_data.json"

def normalizar_data(valor):
    if pd.isna(valor) or valor == "":
        return ""
    try:
        dt = pd.to_datetime(valor)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        texto = str(valor).strip()
        try:
            dt = pd.to_datetime(texto, dayfirst=True)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return texto

def limpar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def main():
    if not ARQUIVO_XLSX.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {ARQUIVO_XLSX}")

    xls = pd.ExcelFile(ARQUIVO_XLSX)
    sheet_name = "PDS_Lancamentos" if "PDS_Lancamentos" in xls.sheet_names else xls.sheet_names[0]

    # A planilha consolidada normalmente tem título nas 3 primeiras linhas e cabeçalho real na 4ª linha.
    # Se a aba vier em outro formato, tenta a primeira linha como cabeçalho.
    try:
        df = pd.read_excel(ARQUIVO_XLSX, sheet_name=sheet_name, header=3)
        if "Data" not in df.columns:
            raise ValueError("Cabecalho padrao nao encontrado")
    except Exception:
        df = pd.read_excel(ARQUIVO_XLSX, sheet_name=sheet_name, header=0)

    # remove linhas totalmente vazias
    df = df.dropna(how="all")

    dados = []
    for _, row in df.iterrows():
        data = normalizar_data(row.get("Data"))
        obra = limpar(row.get("Obra"))
        equipe = limpar(row.get("Equipe"))
        atividade = limpar(row.get("Atividade"))
        pv_inicio = limpar(row.get("PV_Inicio"))
        pv_fim = limpar(row.get("PV_Fim"))
        pv_texto = limpar(row.get("PV_Texto"))

        if not data or not obra or not equipe or not atividade:
            continue

        trecho = ""
        if pv_inicio and pv_fim:
            trecho = f"{pv_inicio}-{pv_fim}"
        elif pv_inicio and not pv_fim:
            trecho = ""

        pv = pv_texto or pv_inicio

        dados.append({
            "data": data,
            "obra": obra,
            "equipe": equipe,
            "atividade": atividade,
            "trecho": trecho,
            "pv": pv
        })

    ARQUIVO_JSON.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"✅ PDS atualizado com sucesso!")
    print(f"Planilha: {ARQUIVO_XLSX}")
    print(f"JSON: {ARQUIVO_JSON}")
    print(f"Registros: {len(dados)}")

if __name__ == "__main__":
    main()
