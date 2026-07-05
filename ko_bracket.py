"""Propagación de equipos en eliminatorias según cuadro FIFA y resultados."""
from __future__ import annotations

import re
import unicodedata

from fifa_schedule_2026 import R32_RESOLVED, build_r32_matches_chronological
from octavos_fixtures import OCTAVOS_CANONICAL

_PLACEHOLDER_RE = re.compile(
    r"^(Ganador|Perdedor)\s",
    re.IGNORECASE,
)


def is_ko_placeholder(name: str) -> bool:
    name = (name or "").strip()
    return bool(name and _PLACEHOLDER_RE.match(name))


def normalize_pair_key(local: str, visitante: str) -> tuple[str, str]:
    def norm(s: str) -> str:
        n = unicodedata.normalize("NFKD", s)
        ascii_name = "".join(c for c in n if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]+", "", ascii_name.lower())

    a, b = norm(local), norm(visitante)
    return tuple(sorted([a, b]))


def build_r32_slot_to_chrono_id() -> dict[int, int]:
    """Mapa slot cuadro R32 (0–15) → id cronológico (73–88)."""
    r32 = build_r32_matches_chronological()
    pair_to_id = {
        normalize_pair_key(m["local"], m["visitante"]): m["id"] for m in r32
    }
    slot_to_id: dict[int, int] = {}
    for slot, (loc, vis) in enumerate(R32_RESOLVED):
        key = normalize_pair_key(loc, vis)
        if key not in pair_to_id:
            raise ValueError(f"Pareja R32 slot {slot} no encontrada: {loc} vs {vis}")
        slot_to_id[slot] = pair_to_id[key]
    return slot_to_id


def resolve_winner(
    local: str,
    visitante: str,
    home: int | None,
    away: int | None,
    clasificado: str | None,
) -> str | None:
    if home is None or away is None:
        return None
    if home > away:
        return local
    if away > home:
        return visitante
    cl = (clasificado or "").strip()
    if cl == "Local":
        return local
    if cl == "Visitante":
        return visitante
    return None


def resolve_loser(
    local: str,
    visitante: str,
    home: int | None,
    away: int | None,
    clasificado: str | None,
) -> str | None:
    winner = resolve_winner(local, visitante, home, away, clasificado)
    if winner is None:
        return None
    return visitante if winner == local else local


def _match_result(
    match_id: int,
    teams: dict[int, tuple[str, str]],
    results: dict[int, dict],
) -> tuple[str | None, str | None]:
    """Devuelve (ganador, perdedor) del partido match_id."""
    if match_id not in teams:
        return None, None
    local, visitante = teams[match_id]
    res = results.get(match_id)
    if not res:
        return None, None
    home, away = res.get("home"), res.get("away")
    cl = res.get("clasificado")
    w = resolve_winner(local, visitante, home, away, cl)
    if w is None:
        return None, None
    loser = visitante if w == local else local
    return w, loser


def compute_resolved_teams(
    results: dict[int, dict],
    calendar: list[dict],
) -> dict[int, tuple[str, str]]:
    """
    Calcula local/visitante reales para partidos KO con placeholders.
    results: {match_id: {home, away, clasificado?}}
    """
    cal_by_id = {m["id"]: m for m in calendar}
    teams: dict[int, tuple[str, str]] = {}

    for m in calendar:
        if not is_ko_placeholder(m["local"]):
            teams[m["id"]] = (m["local"], m["visitante"])

    slot_to_chrono = build_r32_slot_to_chrono_id()
    r32_feeders: list[str | None] = [None] * 16
    for slot in range(16):
        chrono_id = slot_to_chrono[slot]
        local, visitante = R32_RESOLVED[slot]
        teams.setdefault(chrono_id, (local, visitante))
        w, _ = _match_result(chrono_id, teams, results)
        r32_feeders[slot] = w

    for mid, pair in OCTAVOS_CANONICAL.items():
        teams[mid] = pair

    oct_feeders: list[str | None] = [None] * 8
    for i in range(8):
        mid = 89 + i
        w, _ = _match_result(mid, teams, results)
        oct_feeders[i] = w

    for i in range(4):
        mid = 97 + i
        loc, vis = oct_feeders[2 * i], oct_feeders[2 * i + 1]
        if loc and vis:
            teams[mid] = (loc, vis)

    cuartos_feeders: list[str | None] = [None] * 4
    for i in range(4):
        mid = 97 + i
        w, _ = _match_result(mid, teams, results)
        cuartos_feeders[i] = w

    for i in range(2):
        mid = 101 + i
        loc, vis = cuartos_feeders[2 * i], cuartos_feeders[2 * i + 1]
        if loc and vis:
            teams[mid] = (loc, vis)

    semi_winners: list[str | None] = [None, None]
    semi_losers: list[str | None] = [None, None]
    for i in range(2):
        mid = 101 + i
        w, loser = _match_result(mid, teams, results)
        semi_winners[i] = w
        semi_losers[i] = loser

    if semi_winners[0] and semi_winners[1]:
        teams[104] = (semi_winners[0], semi_winners[1])
    if semi_losers[0] and semi_losers[1]:
        teams[103] = (semi_losers[0], semi_losers[1])

    out: dict[int, tuple[str, str]] = {}
    for mid in range(73, 105):
        if mid not in cal_by_id:
            continue
        cal = cal_by_id[mid]
        if mid not in teams:
            continue
        cur = (cal["local"], cal["visitante"])
        new = teams[mid]
        if is_ko_placeholder(cal["local"]):
            out[mid] = new
        elif mid >= 97 and new != cur:
            out[mid] = new
    return out
