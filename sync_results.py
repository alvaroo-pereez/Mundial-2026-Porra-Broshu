"""Sincroniza resultados API-Football → Excel → data.json del dashboard."""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

from api_football import (
    FINISHED_STATUSES,
    extract_result,
    fetch_all_fixtures,
    fixtures_by_id,
    load_calendar,
    load_fixture_map,
    match_is_ko,
)
from config.groups import list_groups, load_group

ROOT = Path(__file__).parent
PT_FIRST_ROW = 4


def read_excel_result(pt, match_id: int) -> tuple:
    row = PT_FIRST_ROW + match_id - 1
    return pt.cell(row, 6).value, pt.cell(row, 7).value, pt.cell(row, 11).value


def write_excel_result(
    pt,
    match_id: int,
    home: int,
    away: int,
    clasificado: str | None,
    is_ko: bool,
) -> bool:
    row = PT_FIRST_ROW + match_id - 1
    old_h, old_a, old_cl = pt.cell(row, 6).value, pt.cell(row, 7).value, pt.cell(row, 11).value
    new_cl = clasificado if is_ko and clasificado else (old_cl if is_ko else "-")

    changed = (
        old_h != home
        or old_a != away
        or (is_ko and old_cl != new_cl)
    )
    if not changed:
        return False

    pt.cell(row, 6, value=home)
    pt.cell(row, 7, value=away)
    if is_ko and clasificado:
        pt.cell(row, 11, value=clasificado)
    return True


def collect_api_updates() -> dict[int, dict]:
    calendar = load_calendar()
    calendar_by_id = {m["id"]: m for m in calendar}
    fixture_map = load_fixture_map()
    all_fixtures = fetch_all_fixtures()
    by_id = fixtures_by_id(all_fixtures)

    updates: dict[int, dict] = {}
    for mid_str, fixture_id in fixture_map.items():
        mid = int(mid_str)
        m = calendar_by_id.get(mid)
        if not m:
            continue
        item = by_id.get(fixture_id)
        if not item:
            continue
        status = item["fixture"]["status"]["short"]
        if status not in FINISHED_STATUSES:
            continue
        parsed = extract_result(item, match_is_ko(m))
        if parsed:
            updates[mid] = parsed
    return updates


def apply_to_excel(excel_path: Path, updates: dict[int, dict], dry_run: bool) -> bool:
    calendar = load_calendar()
    calendar_by_id = {m["id"]: m for m in calendar}
    wb = load_workbook(excel_path)
    pt = wb["Partidos"]
    changed = False

    for mid, result in sorted(updates.items()):
        m = calendar_by_id[mid]
        is_ko = match_is_ko(m)
        old_h, old_a, old_cl = read_excel_result(pt, mid)
        home, away = result["home"], result["away"]
        clasif = result.get("clasificado")

        if dry_run:
            if old_h != home or old_a != away or (is_ko and clasif and old_cl != clasif):
                print(
                    f"  #{mid} {m['local']} vs {m['visitante']}: "
                    f"{old_h}-{old_a} -> {home}-{away}"
                    + (f" clasif={clasif}" if is_ko and clasif else "")
                )
                changed = True
            continue

        if write_excel_result(pt, mid, home, away, clasif, is_ko):
            print(f"  Actualizado #{mid}: {home}-{away}" + (f" ({clasif})" if clasif else ""))
            changed = True

    if changed and not dry_run:
        wb.save(excel_path)
    return changed


def rebuild_dashboards() -> None:
    import build_dashboard

    for group_id in list_groups():
        build_dashboard.main(group_id=group_id, open_browser=False)


def main() -> None:
    dry_run = "--dry-run" in sys.argv
    print("Consultando API-Football...")
    updates = collect_api_updates()
    print(f"Partidos finalizados en API (mapeados): {len(updates)}")

    any_changed = False
    for group_id in list_groups():
        group = load_group(group_id)
        excel_path = group["excel_path"]
        if not excel_path.exists():
            print(f"  Aviso: no existe {excel_path}, se omite")
            continue
        print(f"Grupo {group_id} ({excel_path.name}):")
        if apply_to_excel(excel_path, updates, dry_run):
            any_changed = True

    if any_changed and not dry_run:
        print("Regenerando dashboards...")
        rebuild_dashboards()
        print("Listo.")
    elif not any_changed:
        print("Sin cambios.")
    elif dry_run:
        print("(dry-run: no se escribió nada)")

    sys.exit(0)


if __name__ == "__main__":
    main()
