"""
Lógica de puntuación compartida (Excel + dashboard).
Solo se otorga el tier más alto aplicable; los puntos no son acumulativos.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).parent
SCORING_JSON = ROOT / "config" / "scoring.json"

DEFAULT_CFG: dict[str, Any] = {
    "grupos": {"exacto": 5, "diferencia": 3, "ganador_empate": 2},
    "eliminatorias": {
        "exacto": 10,
        "diferencia": 6,
        "clasificado": 4,
        "empate_exacto_falla_clasificado": 6,
        "empate_diferencia_acierta_clasificado": 6,
        "empate_diferencia_falla_clasificado": 2,
    },
    "bonus_ronda": {"octavos": 15, "cuartos": 15, "semifinal": 15},
    "apuesta_especial": 10,
    "rondas_bonus": {"Octavos": 8, "Cuartos": 4, "Semifinal": 2},
}


def load_scoring_config(path: Path | None = None) -> dict[str, Any]:
    p = path or SCORING_JSON
    if p.exists():
        with p.open(encoding="utf-8") as f:
            return json.load(f)
    return dict(DEFAULT_CFG)


def _sign(h: int, a: int) -> int:
    if h > a:
        return 1
    if h == a:
        return 0
    return -1


def _as_int(v) -> int | None:
    if v is None or v == "":
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _clasif_ok(pred_clasif: str | None, real_clasif: str | None) -> bool:
    pc = (pred_clasif or "").strip()
    rc = (real_clasif or "").strip()
    return bool(pc and rc and pc == rc)


def _calc_ko_points(
    ph: int,
    pa: int,
    rh: int,
    ra: int,
    pred_clasif: str | None,
    real_clasif: str | None,
    ko: dict[str, Any],
) -> int:
    exact = ph == rh and pa == ra
    diff = (ph - pa) == (rh - ra)
    pred_draw = ph == pa
    real_draw = rh == ra
    clas_ok = _clasif_ok(pred_clasif, real_clasif)
    emp_ex_fail = ko.get("empate_exacto_falla_clasificado", ko["diferencia"])
    emp_diff_ok = ko.get("empate_diferencia_acierta_clasificado", ko["diferencia"])
    emp_diff_fail = ko.get("empate_diferencia_falla_clasificado", 2)

    if not real_draw:
        if exact:
            return ko["exacto"]
        if diff:
            return ko["diferencia"]
        if clas_ok:
            return ko["clasificado"]
        return 0

    if exact:
        return ko["exacto"] if clas_ok else emp_ex_fail
    if pred_draw and diff:
        return emp_diff_ok if clas_ok else emp_diff_fail
    if clas_ok:
        return ko["clasificado"]
    return 0


def calc_match_points(
    fase: str,
    pred_h,
    pred_a,
    res_h,
    res_a,
    pred_clasif: str | None,
    real_clasif: str | None,
    cfg: dict[str, Any] | None = None,
) -> int | None:
    """Puntos por partido. None si partido no finalizado o pronóstico incompleto."""
    ph = _as_int(pred_h)
    pa = _as_int(pred_a)
    rh = _as_int(res_h)
    ra = _as_int(res_a)
    if ph is None or pa is None or rh is None or ra is None:
        return None

    c = cfg or load_scoring_config()
    fase = (fase or "").strip()

    if fase == "Grupos":
        g = c["grupos"]
        if ph == rh and pa == ra:
            return g["exacto"]
        if (ph - pa) == (rh - ra):
            return g["diferencia"]
        if _sign(ph, pa) == _sign(rh, ra):
            return g["ganador_empate"]
        return 0

    ko = c["eliminatorias"]
    return _calc_ko_points(ph, pa, rh, ra, pred_clasif, real_clasif, ko)


# Niveles visuales para el dashboard. Mantienen la MISMA prioridad que
# calc_match_points: solo se devuelve el nivel más alto aplicable.
TIER_LABELS = {
    "exact": "Resultado exacto",
    "difference": "Diferencia de goles",
    "winner": "Solo el ganador",
    "clasificado": "Solo el clasificado",
    "miss": "Sin acierto",
    "pending": "Pendiente",
}


def classify_match_tier(
    fase: str,
    pred_h,
    pred_a,
    res_h,
    res_a,
    pred_clasif: str | None,
    real_clasif: str | None,
) -> str:
    """Nivel de acierto de un pronóstico para colorear el dashboard.

    Devuelve uno de: exact, difference, winner, clasificado, miss, pending.
    'pending' = partido no finalizado o pronóstico incompleto.
    """
    ph = _as_int(pred_h)
    pa = _as_int(pred_a)
    rh = _as_int(res_h)
    ra = _as_int(res_a)
    if ph is None or pa is None or rh is None or ra is None:
        return "pending"

    fase = (fase or "").strip()
    if fase == "Grupos":
        if ph == rh and pa == ra:
            return "exact"
        if (ph - pa) == (rh - ra):
            return "difference"
        if _sign(ph, pa) == _sign(rh, ra):
            return "winner"
        return "miss"

    cfg = load_scoring_config()
    ko = cfg["eliminatorias"]
    pts = _calc_ko_points(ph, pa, rh, ra, pred_clasif, real_clasif, ko)
    if pts == ko["exacto"]:
        return "exact"
    emp_diff_fail = ko.get("empate_diferencia_falla_clasificado", 2)
    if pts in (
        ko["diferencia"],
        ko.get("empate_exacto_falla_clasificado", ko["diferencia"]),
        ko.get("empate_diferencia_acierta_clasificado", ko["diferencia"]),
    ):
        return "difference"
    if pts == ko["clasificado"]:
        return "clasificado"
    if pts == emp_diff_fail:
        return "winner"
    return "miss"


def calc_special_points(pred, official, cfg: dict[str, Any] | None = None) -> int:
    c = cfg or load_scoring_config()
    pts = c.get("apuesta_especial", 10)
    if pred is None or official is None:
        return 0
    p = str(pred).strip()
    o = str(official).strip()
    if not p or not o:
        return 0
    return pts if p.lower() == o.lower() else 0


def calc_round_bonus(
    player_correct: int,
    round_finished: int,
    round_total: int,
    bonus_pts: int,
) -> int:
    if round_finished <= 0 or round_finished != round_total:
        return 0
    if player_correct != round_total:
        return 0
    return bonus_pts


def bonus_for_round(
    fase: str,
    correct_clasif_count: int,
    finished_in_round: int,
    cfg: dict[str, Any] | None = None,
) -> int:
    c = cfg or load_scoring_config()
    rondas = c.get("rondas_bonus", {})
    bonus = c.get("bonus_ronda", {})
    key_map = {
        "Octavos": "octavos",
        "Cuartos": "cuartos",
        "Semifinal": "semifinal",
    }
    if fase not in rondas:
        return 0
    n = rondas[fase]
    bkey = key_map.get(fase)
    if not bkey:
        return 0
    return calc_round_bonus(
        correct_clasif_count, finished_in_round, n, bonus.get(bkey, 15)
    )


def run_self_tests() -> None:
    cfg = load_scoring_config()
    cases = [
        ("Grupos", 2, 1, 2, 1, None, None, 5),
        ("Grupos", 3, 1, 2, 0, None, None, 3),
        ("Grupos", 1, 0, 2, 1, None, None, 3),
        ("Grupos", 1, 2, 0, 2, None, None, 2),
        ("Grupos", 1, 1, 2, 2, None, None, 3),
        ("Octavos", 2, 1, 2, 1, "Local", "Local", 10),
        ("Octavos", 3, 1, 2, 0, "Local", "Local", 6),
        ("Octavos", 2, 0, 1, 1, "Local", "Local", 4),
        ("Octavos", 2, 0, 1, 1, "Local", "Visitante", 0),
        ("Dieciseisavos", 1, 1, 1, 1, "Local", "Local", 10),
        ("Dieciseisavos", 1, 1, 1, 1, "Local", "Visitante", 6),
        ("Dieciseisavos", 2, 2, 1, 1, "Local", "Local", 6),
        ("Dieciseisavos", 2, 2, 1, 1, "Visitante", "Local", 2),
        ("Dieciseisavos", 2, 1, 1, 1, "Local", "Local", 4),
    ]
    for i, (fase, ph, pa, rh, ra, pc, rc, expected) in enumerate(cases, 1):
        got = calc_match_points(fase, ph, pa, rh, ra, pc, rc, cfg)
        assert got == expected, f"Caso {i}: esperado {expected}, obtuvo {got}"
    assert calc_special_points("España", "España", cfg) == 10
    assert calc_special_points("España", "Francia", cfg) == 0
    assert calc_round_bonus(8, 8, 8, 15) == 15
    assert calc_round_bonus(7, 8, 8, 15) == 0

    tier_cases = [
        ("Grupos", 2, 1, 2, 1, None, None, "exact"),
        ("Grupos", 3, 1, 2, 0, None, None, "difference"),
        ("Grupos", 1, 0, 2, 1, None, None, "difference"),
        ("Grupos", 1, 2, 0, 2, None, None, "winner"),
        ("Grupos", 2, 0, 0, 1, None, None, "miss"),
        ("Grupos", 1, 1, None, None, None, None, "pending"),
        ("Octavos", 2, 1, 2, 1, "Local", "Local", "exact"),
        ("Octavos", 3, 1, 2, 0, "Local", "Local", "difference"),
        ("Octavos", 2, 0, 1, 1, "Local", "Local", "clasificado"),
        ("Octavos", 2, 0, 1, 1, "Local", "Visitante", "miss"),
        ("Dieciseisavos", 1, 1, 1, 1, "Local", "Local", "exact"),
        ("Dieciseisavos", 2, 2, 1, 1, "Local", "Local", "difference"),
        ("Dieciseisavos", 2, 2, 1, 1, "Visitante", "Local", "winner"),
    ]
    for i, (fase, ph, pa, rh, ra, pc, rc, expected) in enumerate(tier_cases, 1):
        got = classify_match_tier(fase, ph, pa, rh, ra, pc, rc)
        assert got == expected, f"Tier {i}: esperado {expected}, obtuvo {got}"
    print("scoring.py: todos los casos OK")


if __name__ == "__main__":
    run_self_tests()
