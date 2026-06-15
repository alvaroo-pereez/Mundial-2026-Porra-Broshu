"""Genera data/api_fixture_map.json cruzando calendario local con API-Football."""
from __future__ import annotations

import sys
from collections import defaultdict

from api_football import (
    FIXTURE_MAP_PATH,
    load_calendar,
    load_team_mapping,
    match_fixture_to_calendar,
    parse_match_date,
    save_json,
    spanish_to_api,
    fetch_all_fixtures,
)


def main() -> None:
    calendar = load_calendar()
    mapping = load_team_mapping()
    fixtures = fetch_all_fixtures()

    by_date: dict[str, list] = defaultdict(list)
    for item in fixtures:
        date = item["fixture"]["date"][:10]
        by_date[date].append(item)

    fixture_map: dict[str, int] = {}
    unmatched: list[str] = []

    for m in calendar:
        mid = m["id"]
        local = m["local"]
        visitante = m["visitante"]
        if local.startswith("Ganador") or local.startswith("Perdedor"):
            continue

        date_key = parse_match_date(m["fecha"])
        candidates = by_date.get(date_key, [])
        found = None
        for item in candidates:
            if match_fixture_to_calendar(item, local, visitante, mapping):
                found = item
                break

        if found:
            fixture_map[str(mid)] = found["fixture"]["id"]
        else:
            unmatched.append(
                f"  #{mid} {date_key} {local} vs {visitante} "
                f"(API: {spanish_to_api(local, mapping)} vs {spanish_to_api(visitante, mapping)})"
            )

    save_json(FIXTURE_MAP_PATH, fixture_map)
    print(f"Mapeados: {len(fixture_map)} / {len(calendar)} partidos")
    print(f"Guardado: {FIXTURE_MAP_PATH}")

    if unmatched:
        print(f"\nSin mapear ({len(unmatched)}):")
        print("\n".join(unmatched[:20]))
        if len(unmatched) > 20:
            print(f"  ... y {len(unmatched) - 20} más")
        if "--strict" in sys.argv and unmatched:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
