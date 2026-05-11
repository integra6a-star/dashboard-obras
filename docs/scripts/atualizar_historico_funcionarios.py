import json
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

TZ_SP = ZoneInfo("America/Sao_Paulo")
TZ_UTC = ZoneInfo("UTC")

def now_sp():
    return datetime.now(TZ_SP).isoformat()

def now_utc():
    return datetime.now(TZ_UTC).isoformat()

def main():
    # Pastas corretas do seu projeto
    script_dir = Path(__file__).resolve().parent
    root_dir = script_dir.parent
    docs_dir = root_dir / "docs"

    arq_func = docs_dir / "funcionarios.json"
    arq_hist = docs_dir / "funcionarios_historico.json"

    # Validações
    if not arq_func.exists():
        raise FileNotFoundError(f"Não encontrei funcionarios.json em: {arq_func}")

    # Lê funcionarios.json (gerado no passo 2/3)
    with open(arq_func, "r", encoding="utf-8") as f:
        data = json.load(f)

    mes = data.get("mes") or datetime.now(TZ_SP).strftime("%Y-%m")
    resumo = data.get("resumo", {})

    item_mes = {
        "mes": mes,
        "atualizado_em": now_sp(),
        "ativos": int(resumo.get("ativos", 0)),
        "afastados": int(resumo.get("afastados", 0)),
        "total_salarios": float(resumo.get("total_salarios", 0.0)),
        "total_vr": float(resumo.get("total_vr", 0.0)),
        "total_combustivel": float(resumo.get("total_combustivel", 0.0)),
        "total_frota": float(resumo.get("total_frota", 0.0)),
        "total_frota_leve": float(resumo.get("total_frota_leve", 0.0)),
        "total_caminhoes": float(resumo.get("total_caminhoes", 0.0)),
        "total_maquinas": float(resumo.get("total_maquinas", 0.0)),
        "total_outros_ativos": float(resumo.get("total_outros_ativos", 0.0)),
        "custo_mensal_total": float(resumo.get("custo_mensal_total", 0.0)),
        "salario_clt": float(resumo.get("salario_clt", 0.0)),
        "salario_pj": float(resumo.get("salario_pj", 0.0)),
    }

    # Carrega histórico existente (se tiver)
    historico = {"atualizado_em": now_utc(), "series": []}
    if arq_hist.exists():
        try:
            with open(arq_hist, "r", encoding="utf-8") as f:
                historico = json.load(f) or historico
            if "series" not in historico or not isinstance(historico["series"], list):
                historico["series"] = []
        except:
            historico = {"atualizado_em": now_utc(), "series": []}

    # Atualiza/insere o mês
    replaced = False
    for i, it in enumerate(historico["series"]):
        if isinstance(it, dict) and it.get("mes") == mes:
            historico["series"][i] = item_mes
            replaced = True
            break
    if not replaced:
        historico["series"].append(item_mes)

    # Ordena e salva
    historico["series"] = sorted(historico["series"], key=lambda x: x.get("mes", ""))
    historico["atualizado_em"] = now_utc()

    docs_dir.mkdir(parents=True, exist_ok=True)
    with open(arq_hist, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)

    print(f"✅ OK: Histórico atualizado em {arq_hist}")
    print(f"➡️ Mês: {mes} | Ativos: {item_mes['ativos']} | Custo total: {item_mes['custo_mensal_total']}")

if __name__ == "__main__":
    main()