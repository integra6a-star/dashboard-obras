# -*- coding: utf-8 -*-
"""Unifica aliases de São Lucas entre Base Dash e mapa."""

from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PATHS = [ROOT / "de_para_obras.json", ROOT / "docs" / "de_para_obras.json"]


def main() -> None:
    for path in PATHS:
        data = json.loads(path.read_text(encoding="utf-8"))
        target = data.get("RCE SAO LUCAS") or data.get("RCE SÃO LUCAS") or "RCE São Lucas"
        data["rce_comunidade_sao_lucas"] = target
        data["RCE_COMUNIDADE_SAO_LUCAS"] = target
        data["RCE COMUNIDADE SAO LUCAS"] = target
        data["RCE COMUNIDADE SÃO LUCAS"] = target
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"Atualizado: {path}")


if __name__ == "__main__":
    main()
