"""Cliente API-Football compartido (bootstrap + sync)."""
from __future__ import annotations

import json
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import requests

ROOT = Path(__file__).parent
TEAM_MAPPING_PATH = ROOT / "config" / "team_mapping.json"
FIXTURE_MAP_PATH = ROOT / "data" / "api_fixture_map.json"
MATCHES_JSON = ROOT / "data" / "matches_2026.json"

API_BASE = "https://v3.football.api-sports.io"
LEAGUE_ID = 1
SEASON = 2026
FINISHED_STATUSES = frozenset({"FT", "AET", "PEN"})


def get_api_key() -> str:
    key = os.environ.get("API_FOOTBALL_KEY", "").strip()
    if not key:
        raise SystemExit(
            "API_FOOTBALL_KEY no configurada.\n"
            "Regístrate en https://www.api-football.com/ y exporta la clave."
        )
    return key


def load_json(path: Path):
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_team_mapping() -> dict[str, str]:
    return load_json(TEAM_MAPPING_PATH)


def normalize_name(name: str) -> str:
    norm = unicodedata.normalize("NFKD", name)
    ascii_name = "".join(c for c in norm if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", ascii_name.lower())


def spanish_to_api(name: str, mapping: dict[str, str]) -> str:
    if name in mapping:
        return mapping[name]
    return name


def parse_match_date(fecha: str) -> str:
    """DD/MM/YYYY -> YYYY-MM-DD."""
    dt = datetime.strptime(fecha.strip(), "%d/%m/%Y")
    return dt.strftime("%Y-%m-%d")


def fetch_all_fixtures() -> list[dict]:
    key = get_api_key()
    url = f"{API_BASE}/fixtures"
    params = {"league": LEAGUE_ID, "season": SEASON}
    resp = requests.get(
        url,
        headers={"x-apisports-key": key},
        params=params,
        timeout=60,
    )
    resp.raise_for_status()
    payload = resp.json()
    if payload.get("errors"):
        raise RuntimeError(f"API-Football error: {payload['errors']}")
    return payload.get("response", [])


def fixtures_by_id(fixtures: list[dict]) -> dict[int, dict]:
    return {item["fixture"]["id"]: item for item in fixtures}


def team_names_match(api_name: str, expected_api_name: str) -> bool:
    if normalize_name(api_name) == normalize_name(expected_api_name):
        return True
    # Alias frecuentes en la API
    aliases = {
        normalize_name("USA"): {normalize_name("United States"), normalize_name("USA")},
        normalize_name("Ivory Coast"): {
            normalize_name("Côte d'Ivoire"),
            normalize_name("Cote d'Ivoire"),
        },
        normalize_name("DR Congo"): {
            normalize_name("Congo DR"),
            normalize_name("Congo-DR"),
            normalize_name("Democratic Republic of the Congo"),
        },
    }
    a = normalize_name(api_name)
    b = normalize_name(expected_api_name)
    if a == b:
        return True
    for canonical, alts in aliases.items():
        group = alts | {canonical}
        if a in group and b in group:
            return True
    return False


def match_fixture_to_calendar(
    fixture: dict,
    local_es: str,
    visitante_es: str,
    mapping: dict[str, str],
) -> bool:
    home_api = spanish_to_api(local_es, mapping)
    away_api = spanish_to_api(visitante_es, mapping)
    api_home = fixture["teams"]["home"]["name"]
    api_away = fixture["teams"]["away"]["name"]
    return team_names_match(api_home, home_api) and team_names_match(
        api_away, away_api
    )


def extract_result(fixture: dict, is_ko: bool) -> dict | None:
    status = fixture["fixture"]["status"]["short"]
    if status not in FINISHED_STATUSES:
        return None
    goals = fixture["goals"]
    home = goals.get("home")
    away = goals.get("away")
    if home is None or away is None:
        return None
    result: dict = {"home": int(home), "away": int(away)}
    if is_ko:
        winner_home = fixture["teams"]["home"].get("winner")
        if winner_home is True:
            result["clasificado"] = "Local"
        elif winner_home is False:
            result["clasificado"] = "Visitante"
    return result


def load_fixture_map() -> dict[str, int]:
    if not FIXTURE_MAP_PATH.exists():
        raise SystemExit(
            f"No existe {FIXTURE_MAP_PATH}. Ejecuta: py bootstrap_fixture_map.py"
        )
    raw = load_json(FIXTURE_MAP_PATH)
    return {str(k): int(v) for k, v in raw.items()}


def load_calendar() -> list[dict]:
    return load_json(MATCHES_JSON)


def match_is_ko(m: dict) -> bool:
    return m.get("fase", "Grupos") != "Grupos"
