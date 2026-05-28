import argparse
import json
import os
from datetime import datetime
from pathlib import Path
from urllib import error, parse, request


ROOT = Path(__file__).resolve().parents[1]


def load_json(name):
    return json.loads((ROOT / name).read_text(encoding="utf-8-sig"))


def clean_date(value):
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        return None


def clean_timestamp(value):
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).isoformat()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).isoformat()
    except ValueError:
        return None


def clean_month(value):
    if not value:
        return None
    text = str(value).strip()
    if len(text) == 7:
        return f"{text}-01"
    return clean_date(text)


def month_number(value):
    if value in ("", None):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if text.isdigit():
        return int(text)
    months = {
        "jan": 1,
        "fev": 2,
        "mar": 3,
        "abr": 4,
        "mai": 5,
        "jun": 6,
        "jul": 7,
        "ago": 8,
        "set": 9,
        "out": 10,
        "nov": 11,
        "dez": 12,
    }
    return months.get(text[:3].lower())


def number(value):
    if value in ("", None):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


class Supabase:
    def __init__(self, url, key):
        self.url = url.rstrip("/")
        self.key = key

    def request(self, method, table, payload=None, query=None, prefer=None):
        qs = f"?{query}" if query else ""
        req = request.Request(f"{self.url}/rest/v1/{table}{qs}", method=method)
        req.add_header("apikey", self.key)
        req.add_header("Authorization", f"Bearer {self.key}")
        req.add_header("Content-Type", "application/json")
        if prefer:
            req.add_header("Prefer", prefer)
        data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
        try:
            with request.urlopen(req, data=data, timeout=60) as resp:
                body = resp.read().decode("utf-8")
                return json.loads(body) if body else None
        except error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"{method} {table} failed: {exc.code} {body}") from exc

    def insert(self, table, rows, returning=False):
        if not rows:
            return []
        prefer = "return=representation" if returning else "return=minimal"
        return self.request("POST", table, rows, prefer=prefer) or []

    def upsert(self, table, rows, conflict):
        if not rows:
            return []
        query = parse.urlencode({"on_conflict": conflict})
        return self.request("POST", table, rows, query=query, prefer="resolution=merge-duplicates,return=minimal")

    def delete_all(self, table):
        self.request("DELETE", table, query="id=not.is.null", prefer="return=minimal")


def create_batch(db, source, updated_at=None, metadata=None):
    rows = db.insert(
        "import_batches",
        [{
            "source": source,
            "source_updated_at": clean_timestamp(updated_at),
            "metadata": metadata or {},
        }],
        returning=True,
    )
    return rows[0]["id"]


def import_obras(db):
    data = load_json("dados.json")
    batch_id = create_batch(
        db,
        "dados.json",
        data.get("atualizado_em"),
        {
            "fonte_registros": data.get("fonte_registros"),
            "fonte_eap_economias": data.get("fonte_eap_economias"),
            "meses_producao": data.get("meses_producao", []),
        },
    )

    rows = []
    monthly = []
    for item in data.get("registros", []):
        rows.append({
            "batch_id": batch_id,
            "data": clean_date(item.get("Data")),
            "status": item.get("Status"),
            "obra": item.get("Obra") or "Sem obra",
            "bloco": str(item.get("Bloco")) if item.get("Bloco") is not None else None,
            "tipo": item.get("Tipo"),
            "planejado_m": number(item.get("Planejado_m")),
            "executado_m": number(item.get("Executado_m")),
            "pv": number(item.get("PV")),
            "profundidade_m": number(item.get("Profundidade_m")),
            "economias_previstas": number(item.get("Economias_Previstas")),
            "economias_recebidas": number(item.get("Economias_Recebidas")),
            "raw": item,
        })

    inserted = db.insert("obra_registros", rows, returning=True)
    for inserted_row, item in zip(inserted, data.get("registros", [])):
        for month, produced in (item.get("ProducaoMensal") or {}).items():
            monthly.append({
                "obra_registro_id": inserted_row["id"],
                "mes": clean_month(month),
                "produzido_m": number(produced) or 0,
            })
    db.insert("obra_producao_mensal", monthly)

    eap = data.get("eap_producao", {})
    eap_rows = []
    for item in eap.get("mensal", []):
        mes_numero = month_number(item.get("mes"))
        competencia = f"{int(item['ano']):04d}-{mes_numero:02d}-01" if item.get("ano") and mes_numero else None
        eap_rows.append({
            "batch_id": batch_id,
            "ano": item.get("ano"),
            "mes": mes_numero,
            "competencia": competencia,
            "eap": number(item.get("eap")),
            "produzido": number(item.get("produzido")),
            "economias_eap": number(item.get("economias_eap")),
            "economias_recebidas": number(item.get("economias_recebidas")),
            "saldo_mes": number(item.get("saldo_mes")),
            "saldo_economias": number(item.get("saldo_economias")),
            "saldo_acum": number(item.get("saldo_acum")),
            "raw": item,
        })
    db.insert("eap_producao_mensal", eap_rows)
    db.insert("dashboard_snapshots", [{"source": "dados.json", "source_updated_at": data.get("atualizado_em"), "payload": data}])
    return {"obra_registros": len(rows), "obra_producao_mensal": len(monthly), "eap_producao_mensal": len(eap_rows)}


def import_pds(db):
    data = load_json("pds_data.json")
    batch_id = create_batch(db, "pds_data.json", metadata={"total": len(data)})
    rows = [{
        "batch_id": batch_id,
        "data": clean_date(item.get("data")),
        "obra": item.get("obra"),
        "equipe": item.get("equipe"),
        "atividade": item.get("atividade"),
        "trecho": item.get("trecho"),
        "pv": item.get("pv"),
        "raw": item,
    } for item in data]
    db.insert("pds_apontamentos", rows)
    db.insert("dashboard_snapshots", [{"source": "pds_data.json", "payload": data}])
    return {"pds_apontamentos": len(rows)}


def import_eap(db):
    data = load_json("eap_producao.json")
    batch_id = create_batch(db, "eap_producao.json", data.get("atualizado_em"), {"fonte": data.get("fonte"), "aba": data.get("aba")})
    rows = []
    for item in data.get("mensal", []):
        mes_numero = month_number(item.get("mes"))
        competencia = f"{int(item['ano']):04d}-{mes_numero:02d}-01" if item.get("ano") and mes_numero else None
        rows.append({
            "batch_id": batch_id,
            "ano": item.get("ano"),
            "mes": mes_numero,
            "competencia": competencia,
            "eap": number(item.get("eap")),
            "produzido": number(item.get("produzido")),
            "economias_eap": number(item.get("economias_eap")),
            "economias_recebidas": number(item.get("economias_recebidas")),
            "saldo_mes": number(item.get("saldo_mes")),
            "saldo_economias": number(item.get("saldo_economias")),
            "saldo_acum": number(item.get("saldo_acum")),
            "raw": item,
        })
    db.insert("eap_producao_mensal", rows)
    db.insert("dashboard_snapshots", [{"source": "eap_producao.json", "source_updated_at": data.get("atualizado_em"), "payload": data}])
    return {"eap_producao_mensal": len(rows)}


def import_funcionarios(db):
    data = load_json("funcionarios.json")
    batch_id = create_batch(db, "funcionarios.json", data.get("atualizado_em"), {"mes": data.get("mes"), "resumo": data.get("resumo", {})})
    mes = clean_month(data.get("mes"))
    rows = []
    for item in data.get("funcionarios", []):
        rows.append({
            "batch_id": batch_id,
            "mes": mes,
            "setor": item.get("setor"),
            "equipe": item.get("equipe"),
            "responsavel_direto": item.get("responsavel_direto"),
            "matricula": str(item.get("matricula")) if item.get("matricula") is not None else None,
            "nome": item.get("nome"),
            "sexo": item.get("sexo"),
            "funcao": item.get("funcao"),
            "admissao": clean_date(item.get("admissao")),
            "rescisao": clean_date(item.get("rescisao")),
            "status": item.get("status"),
            "regime": item.get("regime"),
            "tipo": item.get("tipo"),
            "salario": number(item.get("salario")),
            "vr": number(item.get("vr")),
            "frota": number(item.get("frota")),
            "combustivel": number(item.get("combustivel")),
            "valor_mensal": number(item.get("valor_mensal")),
            "custo_frota": number(item.get("custo_frota")),
            "categoria_frota": item.get("categoria_frota"),
            "tipo_equipamento": item.get("tipo_equipamento"),
            "custo_total": number(item.get("custo_total")),
            "horario": item.get("horario"),
            "veiculo": item.get("veiculo"),
            "placa": item.get("placa"),
            "raw": item,
        })
    db.insert("funcionarios", rows)
    db.insert("dashboard_snapshots", [{"source": "funcionarios.json", "source_updated_at": data.get("atualizado_em"), "payload": data}])
    return {"funcionarios": len(rows)}


def import_medicao(db):
    data = load_json("medicao.json")
    batch_id = create_batch(db, "medicao.json", data.get("atualizado_em"), {"fonte_excel": data.get("fonte_excel")})
    rows = []
    for tipo in ("medicoes", "amortizacao"):
        for item in data.get(tipo, {}).get("serie", []):
            rows.append({
                "batch_id": batch_id,
                "tipo": tipo,
                "competencia": clean_month(item.get("ym")),
                "descricao": item.get("data"),
                "valor": number(item.get("valor")),
                "acumulado": number(item.get("acumulado")),
                "saldo": number(item.get("saldo")),
                "raw": item,
            })
    db.insert("medicao_series", rows)
    db.insert("dashboard_snapshots", [{"source": "medicao.json", "source_updated_at": data.get("atualizado_em"), "payload": data}])
    return {"medicao_series": len(rows)}


def import_almoxarifado(db):
    data = load_json("almoxarifado.json")
    batch_id = create_batch(db, "almoxarifado.json", data.get("atualizado_em"), {"origem": data.get("origem"), "titulo": data.get("titulo")})
    rows = [{
        "batch_id": batch_id,
        "codigo": item.get("codigo"),
        "descricao": item.get("descricao"),
        "estoque_atual": number(item.get("estoque_atual")),
        "estoque_minimo": number(item.get("estoque_minimo")),
        "entradas": number(item.get("entradas")),
        "saidas": number(item.get("saidas")),
        "quantidade": number(item.get("quantidade")),
        "situacao_planilha": item.get("situacao_planilha"),
        "raw": item,
    } for item in data.get("produtos", [])]
    db.insert("almoxarifado_produtos", rows)
    return {"almoxarifado_produtos": len(rows)}


def import_reclamacoes(db):
    data = load_json("reclamacoes.json")
    batch_id = create_batch(db, "reclamacoes.json", metadata={"total": len(data)})
    rows = []
    for item in data:
        rows.append({
            "id_text": str(item.get("id")),
            "batch_id": batch_id,
            "data": clean_date(item.get("data")),
            "obra": item.get("obra"),
            "morador": item.get("morador"),
            "endereco": item.get("endereco"),
            "tipo_dano": item.get("tipo_dano"),
            "descricao": item.get("descricao"),
            "valor_estimado": number(item.get("valor_estimado")),
            "valor_pago": number(item.get("valor_pago")),
            "status": item.get("status"),
            "responsavel": item.get("responsavel"),
            "observacao": item.get("observacao"),
            "raw": item,
        })
    db.upsert("reclamacoes", rows, "id_text")
    return {"reclamacoes": len(rows)}


def main():
    parser = argparse.ArgumentParser(description="Importa os JSONs do dashboard para Supabase.")
    parser.add_argument("--only", choices=["obras", "eap", "pds", "funcionarios", "medicao", "almoxarifado", "reclamacoes"], action="append")
    parser.add_argument("--write", action="store_true", help="Confirma a gravação no Supabase. Sem isso, o script não altera o banco.")
    args = parser.parse_args()

    if not args.write:
        raise SystemExit(
            "Modo seguro: nenhuma alteração foi feita. "
            "Revise database/schema.sql e rode novamente com --write para importar."
        )

    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise SystemExit("Defina SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY no ambiente.")

    db = Supabase(url, key)
    tasks = {
        "obras": import_obras,
        "eap": import_eap,
        "pds": import_pds,
        "funcionarios": import_funcionarios,
        "medicao": import_medicao,
        "almoxarifado": import_almoxarifado,
        "reclamacoes": import_reclamacoes,
    }
    selected = args.only or list(tasks)
    totals = {}
    for name in selected:
        print(f"Importando {name}...")
        totals.update(tasks[name](db))
    print(json.dumps(totals, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
