"""Audita pronósticos cuartos vs referencia en write_cuartos_predictions."""
from __future__ import annotations

import json
import sys
from pathlib import Path

from write_cuartos_predictions import BROSHU_MANUAL, PAPINENES_MANUAL

ROOT = Path(__file__).parent


def audit_group(data_path: Path, manual: dict, group: str) -> list[str]:
    errors: list[str] = []
    data = json.loads(data_path.read_text(encoding="utf-8"))
    by_name = data["players"]
    for player, expected in manual.items():
        p = by_name.get(player)
        if not p:
            errors.append(f"{group}/{player}: jugador no encontrado")
            continue
        cuartos = {m["id"]: m for m in p["matches"] if m.get("fase") == "Cuartos"}
        for mid, (gh, ga, cl) in expected.items():
            m = cuartos.get(mid)
            if not m:
                errors.append(f"{group}/{player} #{mid}: partido no encontrado")
                continue
            got = (m.get("pred_h"), m.get("pred_a"), m.get("pred_clasificado"))
            if got != (gh, ga, cl):
                errors.append(
                    f"{group}/{player} #{mid} {m['local']} vs {m['visitante']}: "
                    f"got {got}, want {(gh, ga, cl)}"
                )
    return errors


def main() -> None:
    errors = audit_group(ROOT / "output/broshu/data.json", BROSHU_MANUAL, "broshu")
    errors += audit_group(ROOT / "output/papinenes/data.json", PAPINENES_MANUAL, "papinenes")

    broshu = json.loads((ROOT / "output/broshu/data.json").read_text(encoding="utf-8"))
    m98 = next(m for m in broshu["matches"] if m["id"] == 98)
    m99 = next(m for m in broshu["matches"] if m["id"] == 99)
    print(f"Match 98: {m98['local']} vs {m98['visitante']} -> {m98['score']}")
    print(f"Match 99: {m99['local']} vs {m99['visitante']} -> {m99['score']}")

    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  {e}")
        sys.exit(1)
    print("Audit OK: all 14 players cuartos match reference")


if __name__ == "__main__":
    main()
