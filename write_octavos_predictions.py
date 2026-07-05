"""Escribe pronósticos octavos manuales (orden lista → id calendario)."""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook

from config.groups import load_group
from octavos_fixtures import OCTAVOS_LIST_ORDER

PT_FIRST_ROW = 4

# (gh, ga, clasificado) por id calendario
BROSHU_MANUAL: dict[str, dict[int, tuple[int, int, str]]] = {
    "Fer": {
        90: (1, 0, "Local"),
        89: (0, 0, "Local"),
        91: (0, 2, "Visitante"),
        92: (1, 0, "Local"),
        93: (0, 3, "Visitante"),
        94: (2, 0, "Local"),
        95: (3, 1, "Local"),
        96: (1, 3, "Visitante"),
    },
    "Felipe": {
        90: (1, 3, "Visitante"),
        89: (1, 1, "Visitante"),
        91: (2, 2, "Visitante"),
        92: (2, 2, "Visitante"),
        93: (2, 2, "Visitante"),
        94: (2, 1, "Local"),
        95: (3, 1, "Local"),
        96: (0, 1, "Visitante"),
    },
    "Muni": {
        90: (1, 1, "Local"),
        89: (1, 1, "Visitante"),
        91: (1, 1, "Visitante"),
        92: (1, 1, "Local"),
        93: (0, 3, "Visitante"),
        94: (2, 1, "Local"),
        95: (2, 0, "Local"),
        96: (1, 2, "Visitante"),
    },
    "Álvaro": {
        90: (1, 1, "Visitante"),
        89: (0, 2, "Visitante"),
        91: (1, 1, "Local"),
        92: (1, 1, "Local"),
        93: (1, 1, "Visitante"),
        94: (1, 1, "Local"),
        95: (2, 0, "Local"),
        96: (1, 1, "Visitante"),
    },
}


def write_predictions(excel_path: Path, dry_run: bool) -> None:
    wb = load_workbook(excel_path)
    for player, preds in BROSHU_MANUAL.items():
        if player not in wb.sheetnames:
            raise SystemExit(f"Hoja '{player}' no existe en {excel_path.name}")
        ws = wb[player]
        print(f"\n{player}:")
        for mid, (gh, ga, cl) in sorted(preds.items()):
            row = PT_FIRST_ROW + mid - 1
            loc = next((l for m, l, _ in OCTAVOS_LIST_ORDER if m == mid), "?")
            vis = next((v for m, _, v in OCTAVOS_LIST_ORDER if m == mid), "?")
            print(f"  #{mid} {loc} vs {vis}: {gh}-{ga} clasif={cl}")
            if not dry_run:
                ws.cell(row, 6, value=gh)
                ws.cell(row, 7, value=ga)
                ws.cell(row, 8, value=cl)
    if not dry_run:
        wb.save(excel_path)
    wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Pronósticos octavos manuales Broshu")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    excel = load_group("broshu")["excel_path"]
    mode = "DRY-RUN" if args.dry_run else "WRITE"
    print(f"write_octavos_predictions [{mode}] -> {excel.name}")
    write_predictions(excel, args.dry_run)
    print("\nTerminado OK.")


if __name__ == "__main__":
    main()
