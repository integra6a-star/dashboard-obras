import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DATA = "2026-07-22"

REGISTROS = [
    {"data": DATA, "obra": "RCE RUA IGARAÇÚ", "equipe": "Marcio", "atividade": "VCA, PVs e Ligações", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Marcelo", "atividade": "Shaft sondagem do Gás", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "CTS JOÃO CANZI", "equipe": "Goiania Drill", "atividade": "Furo Piloto PV-20 ao PV-21", "trecho": "PV-20 ao PV-21", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Cidão", "atividade": "VCA, PVs e Ligações", "trecho": "PI 36 ao PI 38", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Bruno", "atividade": "Ligações e adequação viela", "trecho": "PI 13 ao PI 16", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Antônio", "atividade": "Shaft PV 43.3", "trecho": "", "pv": "43.3"},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 1", "equipe": "Valter", "atividade": "VCA, PVs e Ligações", "trecho": "PI 71 ao PI 72", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Edvando", "atividade": "VCA, PVs e Ligações", "trecho": "PV 49.1 ao PI 49", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Ricardo", "atividade": "Shaft PV 49.3", "trecho": "", "pv": "49.3"},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Claudinei", "atividade": "VCA, PVs e Ligações", "trecho": "PI 49.2 ao PV 49.3", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Jhonatan", "atividade": "VCA, PVs e Ligações", "trecho": "PI 79 ao 80", "pv": ""},
    {"data": DATA, "obra": "RCE RAULZITO TRECHO 2", "equipe": "Cesar", "atividade": "VCA, PVs e Ligações", "trecho": "PI 83 ao 83A", "pv": ""},
    {"data": DATA, "obra": "RCE SÃO LUCAS", "equipe": "Kaique e Miro", "atividade": "Reposição e acabamentos", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "SERVIÇOS REPOSIÇÃO", "equipe": "Leandro", "atividade": "Reparos Gerais", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "ITI 15 - Henko (Osmar)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "ITI 15 - Henko (Marcio)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "RAULZITO - Henko (Jonathan)", "atividade": "Vacal", "trecho": "", "pv": ""},
    {"data": DATA, "obra": "Vacal", "equipe": "RAULZITO - Henko (Wellington)", "atividade": "Vacal", "trecho": "", "pv": ""},
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
