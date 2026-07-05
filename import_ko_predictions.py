"""
Importa pronósticos de eliminatorias desde porras de amigos al Excel maestro.
Empareja por pareja de equipos, no por ID de partido.
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from config.groups import load_group
from import_r32_predictions import (
    EXTRA_ALIASES,
    SKIP_SHEETS,
    COLS_NEW,
    COLS_OLD,
    _as_int,
    build_pair_index,
    detect_layout,
    find_source_sheet,
    normalize_clasificado,
    pair_key,
    teams_match,
)
from worldcup_data import normalize_name, team_names_match

ROOT = Path(__file__).parent
PT_FIRST_ROW = 4

PHASE_RANGES = {
    "dieciseisavos": (73, 88, "Dieci"),
    "octavos": (89, 96, "Octav"),
    "cuartos": (97, 100, "Cuart"),
    "semifinal": (101, 102, "Semif"),
}

IMPORTS_OCTAVOS: list[tuple[str, str, str, str]] = [
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\PORRA MUNDIAL NACH.xlsx",
        "broshu",
        "Nacho",
        "Nacho",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_Patricio.xlsx",
        "broshu",
        "Patri",
        "Patricio",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\PORRA KIKOTA.xlsx",
        "broshu",
        "Kike",
        "Kike",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_Pepe (1).xlsx",
        "broshu",
        "Pepe",
        "Pepe",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_Correo_Quintero-8AVOS.xlsx",
        "broshu",
        "Quintero",
        "Quintero",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_Correo.xlsx",
        "broshu",
        "Luis",
        "Luis",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_papi.xlsx",
        "papinenes",
        "Papá",
        "Papa",
    ),
    (
        r"c:\Users\Alvaro J Perez Triay\AppData\Local\Temp\Porra_Mundial_2026_Diego.xlsx",
        "papinenes",
        "Diego",
        "Diego",
    ),
]


@dataclass
class KoRow:
    row: int
    file_id: int | None
    local: str
    visit: str
    gh: int | None
    ga: int | None
    clasif_raw: object


@dataclass
class ImportResult:
    label: str
    player: str
    matched: dict[int, tuple[int, int, str | None]] = field(default_factory=dict)
    flags: list[str] = field(default_factory=list)


def load_master_pairs(master_path: Path, first_id: int, last_id: int) -> dict[int, tuple[str, str]]:
    wb = load_workbook(master_path, data_only=True)
    pt = wb["Partidos"]
    pairs: dict[int, tuple[str, str]] = {}
    for mid in range(first_id, last_id + 1):
        row = PT_FIRST_ROW + mid - 1
        loc = pt.cell(row, 4).value
        vis = pt.cell(row, 5).value
        if not loc or not vis:
            continue
        loc_s, vis_s = str(loc).strip(), str(vis).strip()
        if loc_s.startswith("Ganador") or loc_s.startswith("Perdedor"):
            continue
        pairs[mid] = (loc_s, vis_s)
    wb.close()
    return pairs


def extract_ko_rows(ws, cols: dict[str, int], phase_token: str) -> list[KoRow]:
    rows: list[KoRow] = []
    for r in range(PT_FIRST_ROW, ws.max_row + 1):
        fase = ws.cell(r, cols["fase"]).value
        if not fase or phase_token not in str(fase):
            continue
        local = ws.cell(r, cols["local"]).value
        visit = ws.cell(r, cols["visit"]).value
        if not local or not visit:
            continue
        loc_s, vis_s = str(local).strip(), str(visit).strip()
        if loc_s.startswith("Ganador") or loc_s.startswith("Perdedor"):
            continue
        raw_id = ws.cell(r, cols["id"]).value
        file_id = None
        if raw_id is not None:
            try:
                file_id = int(raw_id)
            except (TypeError, ValueError):
                pass
        gh = _as_int(ws.cell(r, cols["gh"]).value)
        ga = _as_int(ws.cell(r, cols["ga"]).value)
        cl = ws.cell(r, cols["cl"]).value
        rows.append(
            KoRow(
                row=r,
                file_id=file_id,
                local=loc_s,
                visit=vis_s,
                gh=gh,
                ga=ga,
                clasif_raw=cl,
            )
        )
    return rows


def parse_source(
    source_path: Path,
    master_pairs: dict[int, tuple[str, str]],
    pair_index: dict[tuple[str, str], int],
    label: str,
    phase_token: str,
) -> ImportResult:
    result = ImportResult(label=label, player="")
    if not source_path.exists():
        result.flags.append(f"ARCHIVO NO ENCONTRADO: {source_path}")
        return result

    wb = load_workbook(source_path, data_only=True)
    ws = find_source_sheet(wb)
    cols = detect_layout(ws)
    layout = "new" if cols is COLS_NEW else "old"
    result.flags.append(f"hoja fuente: {ws.title} (layout {layout})")

    for row in extract_ko_rows(ws, cols, phase_token):
        pk = pair_key(row.local, row.visit)
        mid = pair_index.get(pk)
        if mid is None:
            result.flags.append(
                f"NO MATCH fila {row.row}: {row.local} vs {row.visit} ({row.gh}-{row.ga})"
            )
            continue

        ml, mv = master_pairs[mid]
        swapped = not (teams_match(row.local, ml) and teams_match(row.visit, mv))
        gh, ga = row.gh, row.ga
        if swapped:
            gh, ga = row.ga, row.gh
            result.flags.append(
                f"pareja invertida fila {row.row}: {row.local}-{row.visit} -> master {mid}"
            )

        if row.file_id is not None and row.file_id != mid:
            result.flags.append(f"ID desfasado fila {row.row}: archivo {row.file_id} -> master {mid}")

        if gh is None or ga is None:
            result.flags.append(f"INCOMPLETO master {mid}: goles {gh}-{ga}")
            continue

        cl, cl_flag = normalize_clasificado(row.clasif_raw, ml, mv, gh, ga)
        if cl_flag:
            result.flags.append(f"master {mid}: {cl_flag}")
        if cl is None:
            result.flags.append(f"REVISAR master {mid}: empate sin clasificado ({gh}-{ga})")

        result.matched[mid] = (gh, ga, cl)

    wb.close()
    missing = [m for m in master_pairs if m not in result.matched]
    if missing:
        result.flags.append(f"FALTAN partidos (master resueltos): {missing}")
    return result


def write_ko_to_master(
    master_path: Path,
    player: str,
    predictions: dict[int, tuple[int, int, str | None]],
    dry_run: bool,
) -> None:
    if dry_run:
        return
    wb = load_workbook(master_path)
    if player not in wb.sheetnames:
        wb.close()
        raise SystemExit(f"Hoja '{player}' no existe en {master_path.name}")
    ws = wb[player]
    for mid, (gh, ga, cl) in predictions.items():
        row = PT_FIRST_ROW + mid - 1
        ws.cell(row, 6, value=gh)
        ws.cell(row, 7, value=ga)
        if cl:
            ws.cell(row, 8, value=cl)
    wb.save(master_path)
    wb.close()


def run_imports(phase: str, imports: list[tuple[str, str, str, str]], dry_run: bool) -> int:
    if phase not in PHASE_RANGES:
        raise SystemExit(f"Fase desconocida: {phase}. Usa: {', '.join(PHASE_RANGES)}")
    first_id, last_id, phase_token = PHASE_RANGES[phase]
    total_expected = last_id - first_id + 1
    require_all = phase == "dieciseisavos"

    groups_cache: dict[str, Path] = {}
    master_pairs_cache: dict[str, dict[int, tuple[str, str]]] = {}
    errors = 0

    for source, group_id, player, label in imports:
        if not Path(source).exists():
            print(f"\n[{label}] ERROR: no existe {source}")
            errors += 1
            continue

        if group_id not in groups_cache:
            g = load_group(group_id)
            groups_cache[group_id] = g["excel_path"]
            master_pairs_cache[group_id] = load_master_pairs(
                g["excel_path"], first_id, last_id
            )

        master_path = groups_cache[group_id]
        master_pairs = master_pairs_cache[group_id]
        if not master_pairs:
            print(f"\n[{label}] ERROR: master sin partidos resueltos para {phase}")
            errors += 1
            continue

        pair_index = build_pair_index(master_pairs)
        result = parse_source(Path(source), master_pairs, pair_index, label, phase_token)
        result.player = player

        print(f"\n{'=' * 60}")
        print(f"{label} -> {player} ({master_path.name})")
        print(f"Emparejados: {len(result.matched)}/{len(master_pairs)} (master resueltos)")
        for flag in result.flags:
            print(f"  {flag}")
        for mid in sorted(result.matched):
            gh, ga, cl = result.matched[mid]
            ml, mv = master_pairs[mid]
            cl_s = cl or "?"
            print(f"  {mid:2d} {ml} vs {mv}: {gh}-{ga}  clasif={cl_s}")

        if not result.matched:
            errors += 1
            print("  *** SIN EMPAREJAMIENTOS — no se escribe ***")
            continue

        if require_all and len(result.matched) < total_expected:
            errors += 1
            print(f"  *** INCOMPLETO ({len(result.matched)}/{total_expected}) — no se escribe ***")
            continue

        if any(cl is None for _gh, _ga, cl in result.matched.values()):
            errors += 1
            print("  *** HAY EMPATES SIN CLASIFICADO — no se escribe ***")
            continue

        write_ko_to_master(master_path, player, result.matched, dry_run)
        if dry_run:
            print("  (dry-run: sin escribir)")
        else:
            print(f"  Escrito OK ({len(result.matched)} partidos)")

    return errors


def main() -> None:
    parser = argparse.ArgumentParser(description="Importar pronósticos KO desde porras de amigos")
    parser.add_argument(
        "--phase",
        default="octavos",
        choices=list(PHASE_RANGES),
        help="Fase a importar (default: octavos)",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    imports = IMPORTS_OCTAVOS if args.phase == "octavos" else []
    if not imports:
        print(f"No hay lista IMPORTS definida para fase '{args.phase}'.")
        sys.exit(1)

    mode = "DRY-RUN" if args.dry_run else "IMPORT"
    print(f"Import KO predictions [{mode}] fase={args.phase}")
    errors = run_imports(args.phase, imports, args.dry_run)
    if errors:
        print(f"\nTerminado con {errors} error(es).")
        sys.exit(1)
    print("\nTerminado OK.")


if __name__ == "__main__":
    main()
