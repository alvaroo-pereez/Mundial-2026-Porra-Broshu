"""Cliente openfootball/worldcup.json — resultados Mundial 2026 sin API key."""
from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests

ROOT = Path(__file__).parent
TEAM_MAPPING_PATH = ROOT / "config" / "team_mapping.json"
FIXTURE_MAP_PATH = ROOT / "data" / "api_fixture_map.json"
MATCHES_JSON = ROOT / "data" / "matches_2026.json"

OPENFOOTBALL_URL = (
    "https://raw.githubusercontent.com/openfootball/worldcup.json/master/2026/worldcup.json"
)

_PLACEHOLDER_RE = re.compile(
    r"^(Ganador|Perdedor|Winner|Loser|W\d|L\d|\d+[A-Z]|\d+º|\d+o)",
    re.IGNORECASE,
)


def load_json(path: Path):
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_team_mapping() -> dict[str, str]:
    return load_json(TEAM_MAPPING_PATH)


def load_calendar() -> list[dict]:
    return load_json(MATCHES_JSON)


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


def is_placeholder_name(name: str) -> bool:
    name = (name or "").strip()
    if not name:
        return True
    if _PLACEHOLDER_RE.match(name):
        return True
    if "/" in name or "3rd" in name.lower():
        return True
    return False


def is_calendar_placeholder(local: str) -> bool:
    return local.startswith("Ganador") or local.startswith("Perdedor")


def team_names_match(api_name: str, expected_api_name: str) -> bool:
    if normalize_name(api_name) == normalize_name(expected_api_name):
        return True
    aliases = {
        normalize_name("USA"): {
            normalize_name("United States"),
            normalize_name("USA"),
            normalize_name("Estados Unidos"),
        },
        normalize_name("Ivory Coast"): {
            normalize_name("Côte d'Ivoire"),
            normalize_name("Cote d'Ivoire"),
            normalize_name("Costa de Marfil"),
        },
        normalize_name("DR Congo"): {
            normalize_name("Congo DR"),
            normalize_name("Congo-DR"),
            normalize_name("Democratic Republic of the Congo"),
            normalize_name("RD Congo"),
        },
        normalize_name("Bosnia and Herzegovina"): {
            normalize_name("Bosnia & Herzegovina"),
            normalize_name("Bosnia and Herzegovina"),
            normalize_name("Bosnia y Herzegovina"),
        },
        normalize_name("Curacao"): {
            normalize_name("Curaçao"),
            normalize_name("Curacao"),
        },
        normalize_name("South Korea"): {
            normalize_name("Republic of Korea"),
            normalize_name("Korea Republic"),
            normalize_name("República de Corea"),
        },
        normalize_name("Czech Republic"): {
            normalize_name("República Checa"),
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


def fetch_openfootball_matches() -> list[dict]:
    resp = requests.get(OPENFOOTBALL_URL, timeout=60)
    resp.raise_for_status()
    payload = resp.json()
    matches = payload.get("matches", [])
    if not matches:
        raise RuntimeError("openfootball: JSON sin partidos")
    return matches


def match_openfootball_to_calendar(
    of_item: dict,
    local_es: str,
    visitante_es: str,
    mapping: dict[str, str],
) -> bool:
    home_api = spanish_to_api(local_es, mapping)
    away_api = spanish_to_api(visitante_es, mapping)
    t1 = of_item.get("team1", "")
    t2 = of_item.get("team2", "")
    if is_placeholder_name(t1) or is_placeholder_name(t2):
        return False
    return team_names_match(t1, home_api) and team_names_match(t2, away_api)


def match_is_ko(m: dict) -> bool:
    return m.get("fase", "Grupos") != "Grupos"


def extract_result_from_openfootball(of_item: dict, is_ko: bool) -> dict | None:
    score = of_item.get("score") or {}
    ft = score.get("ft")
    if not ft or len(ft) < 2:
        return None
    home, away = int(ft[0]), int(ft[1])
    result: dict = {"home": home, "away": away}
    if is_ko:
        if home > away:
            result["clasificado"] = "Local"
        elif away > home:
            result["clasificado"] = "Visitante"
    return result


def index_openfootball_matches(matches: list[dict]) -> tuple[dict[int, dict], dict[str, list[dict]]]:
    by_num: dict[int, dict] = {}
    by_date: dict[str, list[dict]] = defaultdict(list)
    for item in matches:
        if "num" in item:
            by_num[int(item["num"])] = item
        date = item.get("date")
        if date:
            by_date[date].append(item)
    return by_num, by_date


def find_openfootball_match(
    cal: dict,
    by_num: dict[int, dict],
    by_date: dict[str, list[dict]],
    mapping: dict[str, str],
) -> dict | None:
    mid = cal["id"]
    local = cal["local"]
    visitante = cal["visitante"]

    if is_calendar_placeholder(local):
        return None

    if mid >= 73 and mid in by_num:
        item = by_num[mid]
        t1, t2 = item.get("team1", ""), item.get("team2", "")
        if is_placeholder_name(t1) or is_placeholder_name(t2):
            return None
        if match_openfootball_to_calendar(item, local, visitante, mapping):
            return item
        return item

    date_key = parse_match_date(cal["fecha"])
    for candidate in by_date.get(date_key, []):
        if match_openfootball_to_calendar(candidate, local, visitante, mapping):
            return candidate
    return None


def collect_finished_updates() -> dict[int, dict]:
    """Devuelve {match_id: {home, away, clasificado?}} para partidos finalizados."""
    calendar = load_calendar()
    mapping = load_team_mapping()
    of_matches = fetch_openfootball_matches()
    by_num, by_date = index_openfootball_matches(of_matches)

    updates: dict[int, dict] = {}
    for cal in calendar:
        mid = cal["id"]
        of_item = find_openfootball_match(cal, by_num, by_date, mapping)
        if not of_item:
            continue
        parsed = extract_result_from_openfootball(of_item, match_is_ko(cal))
        if parsed:
            updates[mid] = parsed
    return updates


def diagnose_mapping() -> tuple[dict[str, int], list[str]]:
    """Cruza calendario local con openfootball; devuelve mapa id->num y no emparejados."""
    calendar = load_calendar()
    mapping = load_team_mapping()
    of_matches = fetch_openfootball_matches()
    by_num, by_date = index_openfootball_matches(of_matches)

    fixture_map: dict[str, int] = {}
    unmatched: list[str] = []

    for cal in calendar:
        mid = cal["id"]
        local = cal["local"]
        visitante = cal["visitante"]
        if is_calendar_placeholder(local):
            continue
        found = find_openfootball_match(cal, by_num, by_date, mapping)
        if found:
            key = found.get("num", mid)
            fixture_map[str(mid)] = int(key)
        else:
            date_key = parse_match_date(cal["fecha"])
            unmatched.append(
                f"  #{mid} {date_key} {local} vs {visitante} "
                f"(OF: {spanish_to_api(local, mapping)} vs {spanish_to_api(visitante, mapping)})"
            )
    return fixture_map, unmatched
