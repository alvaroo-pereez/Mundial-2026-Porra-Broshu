"""Audita pronósticos semifinal/final vs referencia y resultados."""
from __future__ import annotations

import json
import sys
from pathlib import Path

from write_semifinal_predictions import (
    BROSHU_SEMIFINAL,
    FINAL_TERCER,
    PAPINENES_SEMIFINAL,
)

ROOT = Path(__file__).parent


def audit_preds(
    data_path: Path,
    semi_ref: dict[str, dict[int, tuple[int, int, str]]],
    final_ref: dict[str, dict[int, tuple[int, int, str]]],
    clear_semifinal: list[str],
    group: str,
) -> list[str]:
    errors: list[str] = []
    data = json.loads(data_path.read_text(encoding="utf-8"))
    players = data["players"]

    for player in clear_semifinal:
        if player not in players:
            errors.append(f"{group}/{player}: jugador no encontrado")
            continue
        for mid in (101, 102):
            m = _match(players[player], mid)
            if m.get("pred_h") is not None or m.get("pred_a") is not None:
                errors.append(f"{group}/{player} #{mid}: debería estar vacío")

    for player, expected in semi_ref.items():
        if player not in players:
            errors.append(f"{group}/{player}: jugador no encontrado")
            continue
        for mid, (gh, ga, cl) in expected.items():
            m = _match(players[player], mid)
            got = (m.get("pred_h"), m.get("pred_a"), m.get("pred_clasificado"))
            if got != (gh, ga, cl):
                errors.append(
                    f"{group}/{player} #{mid} {m['local']} vs {m['visitante']}: "
                    f"got {got}, want {(gh, ga, cl)}"
                )

    all_in_group = set(semi_ref) | set(final_ref) | set(clear_semifinal)
    for player in all_in_group:
        if player not in players:
            continue
        for mid in (103, 104):
            m = _match(players[player], mid)
            exp = final_ref.get(player, {}).get(mid)
            if exp is None:
                if m.get("pred_h") is not None or m.get("pred_a") is not None:
                    errors.append(f"{group}/{player} #{mid}: debería estar vacío")
            else:
                gh, ga, cl = exp
                got = (m.get("pred_h"), m.get("pred_a"), m.get("pred_clasificado"))
                if got != (gh, ga, cl):
                    errors.append(
                        f"{group}/{player} #{mid} {m['local']} vs {m['visitante']}: "
                        f"got {got}, want {(gh, ga, cl)}"
                    )
    return errors


def _match(player: dict, mid: int) -> dict:
    for m in player["matches"]:
        if m["id"] == mid:
            return m
    raise KeyError(mid)


def audit_results(data_path: Path) -> list[str]:
    errors: list[str] = []
    data = json.loads(data_path.read_text(encoding="utf-8"))
    by_id = {m["id"]: m for m in data["matches"]}

    checks = [
        (98, "España", "Bélgica", "2-1"),
        (99, "Noruega", "Inglaterra", "1-1"),
        (101, "Francia", "España", "0-2"),
        (102, "Inglaterra", "Argentina", "1-2"),
        (103, "Francia", "Inglaterra", None),
        (104, "España", "Argentina", None),
    ]
    for mid, loc, vis, score in checks:
        m = by_id[mid]
        if m["local"] != loc or m["visitante"] != vis:
            errors.append(f"#{mid} teams: {m['local']} vs {m['visitante']}, want {loc} vs {vis}")
        if score and m.get("score") != score:
            errors.append(f"#{mid} score: {m.get('score')}, want {score}")
    return errors


def main() -> None:
    errors = audit_preds(
        ROOT / "output/broshu/data.json",
        BROSHU_SEMIFINAL,
        FINAL_TERCER,
        ["Felipe"],
        "broshu",
    )
    errors += audit_preds(
        ROOT / "output/papinenes/data.json",
        PAPINENES_SEMIFINAL,
        {},
        [],
        "papinenes",
    )
    errors += audit_results(ROOT / "output/broshu/data.json")

    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  {e}")
        sys.exit(1)
    print("Audit OK: semifinales, final/tercer y resultados correctos")


if __name__ == "__main__":
    main()
