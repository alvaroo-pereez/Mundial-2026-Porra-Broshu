"""Comprueba que el pipeline de automatización está listo para ejecutarse."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

from config.groups import list_groups, load_group
from worldcup_data import fetch_openfootball_matches

ROOT = Path(__file__).parent
PT_FIRST_ROW = 4


def _log(level: str, message: str) -> None:
    print(f"[{level}] {message}")


def check_excel_files() -> list[str]:
    errors: list[str] = []
    for group_id in list_groups():
        group = load_group(group_id)
        excel_path = group["excel_path"]
        if not excel_path.exists():
            errors.append(f"Falta Excel del grupo {group_id}: {excel_path}")
            continue
        try:
            wb = load_workbook(excel_path, read_only=True, data_only=True)
            if "Partidos" not in wb.sheetnames:
                errors.append(f"{excel_path.name}: no tiene hoja 'Partidos'")
            wb.close()
        except Exception as exc:
            errors.append(f"No se puede leer {excel_path.name}: {exc}")
    return errors


def check_openfootball(test_fetch: bool) -> list[str]:
    errors: list[str] = []
    if not test_fetch:
        return errors
    try:
        matches = fetch_openfootball_matches()
        finished = sum(1 for m in matches if (m.get("score") or {}).get("ft"))
        _log("OK", f"openfootball responde ({len(matches)} partidos, {finished} con resultado)")
        if len(matches) < 100:
            errors.append(f"openfootball devolvió solo {len(matches)} partidos (esperados ~104)")
    except Exception as exc:
        errors.append(f"openfootball no responde: {exc}")
    return errors


def check_predictions_sample() -> None:
    """Informa si hay pronósticos rellenados (sanity check, no bloquea)."""
    for group_id in list_groups():
        group = load_group(group_id)
        excel_path = group["excel_path"]
        if not excel_path.exists():
            continue
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        players = group["players"]
        filled = 0
        for name in players:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            for row in range(PT_FIRST_ROW, PT_FIRST_ROW + 104):
                home = ws.cell(row, 4).value
                away = ws.cell(row, 5).value
                if home is not None and away is not None:
                    filled += 1
        wb.close()
        _log("OK", f"Grupo {group_id}: pronósticos en hojas de jugadores (muestras filas: {filled})")


def main() -> None:
    parser = argparse.ArgumentParser(description="Verifica el pipeline de automatización")
    parser.add_argument(
        "--skip-api",
        action="store_true",
        help="No llama a openfootball (solo comprueba Excel)",
    )
    args = parser.parse_args()

    _log("INFO", "Verificando automatización porra Mundial 2026...")
    errors: list[str] = []
    errors.extend(check_excel_files())
    errors.extend(check_openfootball(test_fetch=not args.skip_api))

    try:
        check_predictions_sample()
    except Exception as exc:
        _log("WARN", f"No se pudo comprobar pronósticos: {exc}")

    if errors:
        for err in errors:
            _log("ERROR", err)
        sys.exit(1)

    _log("INFO", "Todo listo para automatizar.")


if __name__ == "__main__":
    main()
