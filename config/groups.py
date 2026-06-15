"""Carga de configuración por grupo (Broshu, Papinenes, ...)."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GROUPS_DIR = Path(__file__).resolve().parent / "groups"


def list_groups() -> list[str]:
    return sorted(p.stem for p in GROUPS_DIR.glob("*.json"))


def load_group(group_id: str) -> dict:
    path = GROUPS_DIR / f"{group_id}.json"
    if not path.exists():
        known = ", ".join(list_groups()) or "(ninguno)"
        raise SystemExit(f"Grupo desconocido: {group_id}. Disponibles: {known}")
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    data["id"] = data.get("id", group_id)
    data["output_dir"] = ROOT / "output"
    data["excel_path"] = data["output_dir"] / data["excel"]
    data["dashboard_path"] = data["output_dir"] / data["dashboard"]
    data["assets_path"] = data["output_dir"] / data["assets_dir"]
    data["photos_path"] = ROOT / data["photos_src"]
    data["photo_overrides"] = data.get("photo_overrides", {})
    return data


def resolve_group_ids(argv: list[str]) -> list[str]:
    if "--group" not in argv:
        return ["broshu"]
    idx = argv.index("--group")
    if idx + 1 >= len(argv):
        raise SystemExit("Falta el id de grupo tras --group (broshu, papinenes o all).")
    value = argv[idx + 1].lower()
    if value == "all":
        return list_groups()
    return [value]
