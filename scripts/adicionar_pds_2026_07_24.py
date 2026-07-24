import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DATA = "2026-07-24"

REGISTROS = [
    {"data": DATA, "obra": "RCE RUA IGARAÇÚ", "equipe": "Cesar", "atividade": "VCA, PVs e Ligações", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Marcio", "atividade": "REPOSIÇÃO FÁBRICA", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Marcelo", "atividade": "Transformação PV-21", "trecho": "", "pv": "21"},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Antônio", "atividade": "Transformação PV 20.A", "trecho": "", "pv": "20.A"},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Claudinei", "atividade": "Transformação PV 20", "trecho": "", "pv": "20"},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Cidão", "atividade": "VCA, PVs e Ligações", "trecho": "PI 36 ao PI 38", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Bruno", "atividade": "Ligações e adequação viela", "trecho": "PI 13 ao PI 16", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Valter", "atividade": "VCA, PVs e Ligações", "trecho": "PI 83 ao PI 84", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Edvando", "atividade": "Acabamento nos PVs", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Jhonatan", "atividade": "VCA, PVs e Ligações", "trecho": "PI 81 ao PI 82", "pv": ""},
    {"data": DATA, "obra": "RCE SÃO LUCAS", "equipe": "Kaique e Miro", "atividade": "Reposição e acabamentos", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "SERVIÇOS REPOSIÇÃO", "equipe": "Leandro", "atividade": "Reparos Gerais", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "ITI 15 - Henko (Osmar)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "ITI 15 - Henko (Marcio)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "CTs João Canzi - Henko (Jonathan)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "CTs João Canzi - Henko (Wellington)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Guindauto", "equipe": "Nascimento", "atividade": "Apoio às equipes", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Guindauto", "equipe": "Luiz", "atividade": "Apoio às equipes", "trecho": "", "pv": ""},
]


def atualizar(path: Path) -> int:
    dados = json.loads(path.read_text(encoding="utf-8"))
    dados = [item for item in dados if item.get("data") != DATA]
    dados.extend(REGISTROS)
    path.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(dados)


def main():
    total_docs = atualizar(ROOT / "docs" / "pds_data.json")
    total_root = atualizar(ROOT / "pds_data.json")
    print(f"PDS {DATA} atualizado: {len(REGISTROS)} registro(s).")
    print(f"docs/pds_data.json: {total_docs} registro(s).")
    print(f"pds_data.json: {total_root} registro(s).")


if __name__ == "__main__":
    main()
