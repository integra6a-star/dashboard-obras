# -*- coding: utf-8 -*-
"""Importa somente as datas faltantes do PDS de junho.

O importador padrao traz apenas a data mais recente do Word. Este ajuste
recompoe o intervalo que ficou ausente no seletor: 03/06 a 08/06 e 11/06.
"""

from pathlib import Path

import importar_pds_word as imp


DOCX = Path.home() / "Desktop" / "PDS" / "PDS-Junho.docx"
TARGET_DATES = {
    "2026-06-03",
    "2026-06-04",
    "2026-06-05",
    "2026-06-06",
    "2026-06-08",
    "2026-06-11",
}


def main() -> None:
    rows = [
        row for row in imp.parse_docx(DOCX)
        if row["Data"].strftime("%Y-%m-%d") in TARGET_DATES
    ]
    inserted, removed, total = imp.import_rows(rows)
    by_date = {}
    for row in rows:
        key = row["Data"].strftime("%Y-%m-%d")
        by_date[key] = by_date.get(key, 0) + 1
    print(f"Datas faltantes importadas: {total} linhas")
    print(f"Removidas para reimportacao: {removed} | Inseridas: {inserted}")
    print(", ".join(f"{date}={by_date[date]}" for date in sorted(by_date)))


if __name__ == "__main__":
    main()
