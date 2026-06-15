"""Diagnóstico: cruza calendario local con openfootball/worldcup.json."""
from __future__ import annotations

import sys

from worldcup_data import FIXTURE_MAP_PATH, diagnose_mapping, save_json


def main() -> None:
    fixture_map, unmatched = diagnose_mapping()
    save_json(FIXTURE_MAP_PATH, fixture_map)
    print(f"Emparejados: {len(fixture_map)} partidos con nombres reales")
    print(f"Informe opcional guardado: {FIXTURE_MAP_PATH}")

    if unmatched:
        print(f"\nSin emparejar ({len(unmatched)}):")
        print("\n".join(unmatched[:20]))
        if len(unmatched) > 20:
            print(f"  ... y {len(unmatched) - 20} más")
        if "--strict" in sys.argv and unmatched:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
