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
    return teams_orientation(of_item, local_es, visitante_es, mapping) is not None


def teams_orientation(
    of_item: dict,
    local_es: str,
    visitante_es: str,
    mapping: dict[str, str],
) -> str | None:
    """Devuelve 'normal' si team1=local, 'swapped' si team1=visitante, None si no coincide."""
    home_api = spanish_to_api(local_es, mapping)
    away_api = spanish_to_api(visitante_es, mapping)
    t1 = of_item.get("team1", "")
    t2 = of_item.get("team2", "")
    if is_placeholder_name(t1) or is_placeholder_name(t2):
        return None
    if team_names_match(t1, home_api) and team_names_match(t2, away_api):
        return "normal"
    if team_names_match(t1, away_api) and team_names_match(t2, home_api):
        return "swapped"
    return None


def match_is_ko(m: dict) -> bool:
    return m.get("fase", "Grupos") != "Grupos"


def _clasificado_from_t1_t2(
    t1_score: int,
    t2_score: int,
    orientation: str,
) -> str | None:
    if t1_score == t2_score:
        return None
    t1_wins = t1_score > t2_score
    if orientation == "normal":
        return "Local" if t1_wins else "Visitante"
    return "Visitante" if t1_wins else "Local"


def extract_result_from_openfootball(
    of_item: dict,
    is_ko: bool,
    orientation: str = "normal",
) -> dict | None:
    """Goles a 90' (score.ft). Clasificado KO por marcador o penaltis (score.p)."""
    score = of_item.get("score") or {}
    ft = score.get("ft")
    if not ft or len(ft) < 2:
        return None
    t1, t2 = int(ft[0]), int(ft[1])
    if orientation == "swapped":
        home, away = t2, t1
    else:
        home, away = t1, t2

    result: dict = {"home": home, "away": away}
    if not is_ko:
        return result

    if home > away:
        result["clasificado"] = "Local"
    elif away > home:
        result["clasificado"] = "Visitante"
    else:
        pens = score.get("p")
        if pens and len(pens) >= 2:
            cl = _clasificado_from_t1_t2(int(pens[0]), int(pens[1]), orientation)
            if cl:
                result["clasificado"] = cl
        else:
            et = score.get("et")
            if et and len(et) >= 2:
                cl = _clasificado_from_t1_t2(int(et[0]), int(et[1]), orientation)
                if cl:
                    result["clasificado"] = cl
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
    local = cal["local"]
    visitante = cal["visitante"]

    if is_calendar_placeholder(local):
        return None

    date_key = parse_match_date(cal["fecha"])
    for candidate in by_date.get(date_key, []):
        if teams_orientation(candidate, local, visitante, mapping) is not None:
            return candidate

    # Respaldo: ids 1–72 suelen coincidir num con openfootball si la fecha falla
    mid = cal["id"]
    if mid in by_num:
        item = by_num[mid]
        if teams_orientation(item, local, visitante, mapping) is not None:
            return item

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
        orientation = teams_orientation(
            of_item, cal["local"], cal["visitante"], mapping
        )
        if not orientation:
            continue
        parsed = extract_result_from_openfootball(
            of_item, match_is_ko(cal), orientation
        )
        if parsed:
            updates[mid] = parsed
    return updates


def diagnose_mapping() -> tuple[dict[str, int], list[str], list[str]]:
    """Cruza calendario local con openfootball; mapa id->num, no emparejados y avisos."""
    calendar = load_calendar()
    mapping = load_team_mapping()
    of_matches = fetch_openfootball_matches()
    by_num, by_date = index_openfootball_matches(of_matches)

    fixture_map: dict[str, int] = {}
    unmatched: list[str] = []
    warnings: list[str] = []

    for cal in calendar:
        mid = cal["id"]
        local = cal["local"]
        visitante = cal["visitante"]
        if is_calendar_placeholder(local):
            continue
        found = find_openfootball_match(cal, by_num, by_date, mapping)
        if found:
            of_num = int(found.get("num", mid))
            fixture_map[str(mid)] = of_num
            if of_num != mid:
                warnings.append(
                    f"  #{mid} calendario -> OF num {of_num} "
                    f"({local} vs {visitante}; emparejado por equipos)"
                )
        else:
            date_key = parse_match_date(cal["fecha"])
            unmatched.append(
                f"  #{mid} {date_key} {local} vs {visitante} "
                f"(OF: {spanish_to_api(local, mapping)} vs "
                f"{spanish_to_api(visitante, mapping)})"
            )
    return fixture_map, unmatched, warnings


def run_self_tests() -> None:
    mapping = load_team_mapping()

    brasil_japon = {
        "num": 76,
        "date": "2026-06-29",
        "team1": "Brazil",
        "team2": "Japan",
        "score": {"ft": [2, 1], "ht": [0, 1]},
    }
    cal_bj = {
        "id": 74,
        "fecha": "29/06/2026",
        "local": "Brasil",
        "visitante": "Japón",
        "fase": "Dieciseisavos",
    }
    wrong_num = {
        "num": 74,
        "date": "2026-06-29",
        "team1": "Germany",
        "team2": "Paraguay",
        "score": {"ft": [1, 1], "p": [3, 4], "ht": [0, 1]},
    }
    by_num, by_date = index_openfootball_matches([brasil_japon, wrong_num])
    found = find_openfootball_match(cal_bj, by_num, by_date, mapping)
    assert found is brasil_japon, "Brasil-Japón debe emparejar OF num 76, no 74"
    assert teams_orientation(found, cal_bj["local"], cal_bj["visitante"], mapping) == "normal"

    orient = teams_orientation(wrong_num, "Alemania", "Paraguay", mapping)
    res = extract_result_from_openfootball(wrong_num, True, orient)
    assert res == {"home": 1, "away": 1, "clasificado": "Visitante"}, res

    holanda = {
        "num": 75,
        "team1": "Netherlands",
        "team2": "Morocco",
        "score": {"ft": [1, 1], "p": [2, 3]},
    }
    res_h = extract_result_from_openfootball(
        holanda, True, teams_orientation(holanda, "Países Bajos", "Marruecos", mapping)
    )
    assert res_h == {"home": 1, "away": 1, "clasificado": "Visitante"}, res_h

    swapped = {
        "team1": "Japan",
        "team2": "Brazil",
        "score": {"ft": [1, 2]},
    }
    res_s = extract_result_from_openfootball(swapped, True, "swapped")
    assert res_s == {"home": 2, "away": 1, "clasificado": "Local"}, res_s

    draw_no_pens = {"score": {"ft": [1, 1]}}
    res_d = extract_result_from_openfootball(draw_no_pens, True, "normal")
    assert res_d == {"home": 1, "away": 1}, res_d

    bel_sen = {
        "num": 82,
        "date": "2026-07-01",
        "team1": "Belgium",
        "team2": "Senegal",
        "score": {"et": [3, 2], "ft": [2, 2], "ht": [0, 1]},
    }
    cal_bs = {
        "id": 81,
        "fecha": "01/07/2026",
        "local": "Bélgica",
        "visitante": "Senegal",
        "fase": "Dieciseisavos",
    }
    by_num_bs, by_date_bs = index_openfootball_matches([bel_sen])
    found_bs = find_openfootball_match(cal_bs, by_num_bs, by_date_bs, mapping)
    assert found_bs is bel_sen, "Bélgica-Senegal debe emparejar por fecha+equipos"
    orient_bs = teams_orientation(found_bs, cal_bs["local"], cal_bs["visitante"], mapping)
    res_bs = extract_result_from_openfootball(found_bs, True, orient_bs)
    assert res_bs == {"home": 2, "away": 2, "clasificado": "Local"}, res_bs


if __name__ == "__main__":
    run_self_tests()
    print("worldcup_data: tests OK")
