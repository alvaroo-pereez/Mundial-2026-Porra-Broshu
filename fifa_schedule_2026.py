"""Calendario oficial FIFA Mundial 2026 (post-sorteo diciembre 2025)."""

from __future__ import annotations

GROUPS: dict[str, list[str]] = {
    "A": ["México", "Sudáfrica", "República de Corea", "República Checa"],
    "B": ["Canadá", "Catar", "Suiza", "Bosnia y Herzegovina"],
    "C": ["Brasil", "Marruecos", "Haití", "Escocia"],
    "D": ["Estados Unidos", "Paraguay", "Australia", "Turquía"],
    "E": ["Alemania", "Curazao", "Costa de Marfil", "Ecuador"],
    "F": ["Países Bajos", "Japón", "Túnez", "Suecia"],
    "G": ["Bélgica", "Egipto", "Irán", "Nueva Zelanda"],
    "H": ["España", "Cabo Verde", "Arabia Saudí", "Uruguay"],
    "I": ["Francia", "Senegal", "Irak", "Noruega"],
    "J": ["Argentina", "Argelia", "Austria", "Jordania"],
    "K": ["Portugal", "RD Congo", "Uzbekistán", "Colombia"],
    "L": ["Inglaterra", "Croacia", "Ghana", "Panamá"],
}

# Orden de la porra (nº partido 1-72): local, visitante por fila del Excel oficial.
PORRA_GROUP_FIXTURES: list[tuple[str, str]] = [
    ("México", "Sudáfrica"),
    ("República de Corea", "República Checa"),
    ("Canadá", "Bosnia y Herzegovina"),
    ("Estados Unidos", "Paraguay"),
    ("Catar", "Suiza"),
    ("Brasil", "Marruecos"),
    ("Haití", "Escocia"),
    ("Australia", "Turquía"),
    ("Alemania", "Curazao"),
    ("Países Bajos", "Japón"),
    ("Costa de Marfil", "Ecuador"),
    ("Suecia", "Túnez"),
    ("España", "Cabo Verde"),
    ("Bélgica", "Egipto"),
    ("Arabia Saudí", "Uruguay"),
    ("Irán", "Nueva Zelanda"),
    ("Francia", "Senegal"),
    ("Irak", "Noruega"),
    ("Argentina", "Argelia"),
    ("Austria", "Jordania"),
    ("Portugal", "RD Congo"),
    ("Inglaterra", "Croacia"),
    ("Ghana", "Panamá"),
    ("Uzbekistán", "Colombia"),
    ("República Checa", "Sudáfrica"),
    ("Suiza", "Bosnia y Herzegovina"),
    ("Canadá", "Catar"),
    ("México", "República de Corea"),
    ("Estados Unidos", "Australia"),
    ("Escocia", "Marruecos"),
    ("Brasil", "Haití"),
    ("Turquía", "Paraguay"),
    ("Países Bajos", "Suecia"),
    ("Alemania", "Costa de Marfil"),
    ("Ecuador", "Curazao"),
    ("Túnez", "Japón"),
    ("España", "Arabia Saudí"),
    ("Bélgica", "Irán"),
    ("Uruguay", "Cabo Verde"),
    ("Nueva Zelanda", "Egipto"),
    ("Argentina", "Austria"),
    ("Francia", "Irak"),
    ("Noruega", "Senegal"),
    ("Jordania", "Argelia"),
    ("Portugal", "Uzbekistán"),
    ("Inglaterra", "Ghana"),
    ("Panamá", "Croacia"),
    ("Colombia", "RD Congo"),
    ("Suiza", "Canadá"),
    ("Bosnia y Herzegovina", "Catar"),
    ("Escocia", "Brasil"),
    ("Marruecos", "Haití"),
    ("República Checa", "México"),
    ("Sudáfrica", "República de Corea"),
    ("Curazao", "Costa de Marfil"),
    ("Ecuador", "Alemania"),
    ("Japón", "Suecia"),
    ("Túnez", "Países Bajos"),
    ("Turquía", "Estados Unidos"),
    ("Paraguay", "Australia"),
    ("Noruega", "Francia"),
    ("Senegal", "Irak"),
    ("Cabo Verde", "Arabia Saudí"),
    ("Uruguay", "España"),
    ("Egipto", "Irán"),
    ("Nueva Zelanda", "Bélgica"),
    ("Panamá", "Inglaterra"),
    ("Croacia", "Ghana"),
    ("Colombia", "Portugal"),
    ("RD Congo", "Uzbekistán"),
    ("Argelia", "Austria"),
    ("Jordania", "Argentina"),
]

# Calendario FIFA por pareja (local, visitante) -> fecha, hora local, grupo.
FIFA_PAIR_META: dict[tuple[str, str], dict[str, str]] = {
    ("México", "Sudáfrica"): {"fecha": "11/06/2026", "hora": "13:00", "grupo": "A"},
    ("República de Corea", "República Checa"): {"fecha": "11/06/2026", "hora": "20:00", "grupo": "A"},
    ("Canadá", "Bosnia y Herzegovina"): {"fecha": "12/06/2026", "hora": "15:00", "grupo": "B"},
    ("Estados Unidos", "Paraguay"): {"fecha": "12/06/2026", "hora": "18:00", "grupo": "D"},
    ("Haití", "Escocia"): {"fecha": "13/06/2026", "hora": "21:00", "grupo": "C"},
    ("Australia", "Turquía"): {"fecha": "13/06/2026", "hora": "21:00", "grupo": "D"},
    ("Brasil", "Marruecos"): {"fecha": "13/06/2026", "hora": "18:00", "grupo": "C"},
    ("Catar", "Suiza"): {"fecha": "13/06/2026", "hora": "12:00", "grupo": "B"},
    ("Costa de Marfil", "Ecuador"): {"fecha": "14/06/2026", "hora": "19:00", "grupo": "E"},
    ("Alemania", "Curazao"): {"fecha": "14/06/2026", "hora": "12:00", "grupo": "E"},
    ("Países Bajos", "Japón"): {"fecha": "14/06/2026", "hora": "15:00", "grupo": "F"},
    ("Suecia", "Túnez"): {"fecha": "14/06/2026", "hora": "20:00", "grupo": "F"},
    ("Arabia Saudí", "Uruguay"): {"fecha": "15/06/2026", "hora": "18:00", "grupo": "H"},
    ("España", "Cabo Verde"): {"fecha": "15/06/2026", "hora": "12:00", "grupo": "H"},
    ("Irán", "Nueva Zelanda"): {"fecha": "15/06/2026", "hora": "18:00", "grupo": "G"},
    ("Bélgica", "Egipto"): {"fecha": "15/06/2026", "hora": "12:00", "grupo": "G"},
    ("Francia", "Senegal"): {"fecha": "16/06/2026", "hora": "15:00", "grupo": "I"},
    ("Irak", "Noruega"): {"fecha": "16/06/2026", "hora": "18:00", "grupo": "I"},
    ("Argentina", "Argelia"): {"fecha": "16/06/2026", "hora": "20:00", "grupo": "J"},
    ("Austria", "Jordania"): {"fecha": "16/06/2026", "hora": "21:00", "grupo": "J"},
    ("Ghana", "Panamá"): {"fecha": "17/06/2026", "hora": "19:00", "grupo": "L"},
    ("Inglaterra", "Croacia"): {"fecha": "17/06/2026", "hora": "15:00", "grupo": "L"},
    ("Portugal", "RD Congo"): {"fecha": "17/06/2026", "hora": "12:00", "grupo": "K"},
    ("Uzbekistán", "Colombia"): {"fecha": "17/06/2026", "hora": "20:00", "grupo": "K"},
    ("República Checa", "Sudáfrica"): {"fecha": "18/06/2026", "hora": "12:00", "grupo": "A"},
    ("Suiza", "Bosnia y Herzegovina"): {"fecha": "18/06/2026", "hora": "12:00", "grupo": "B"},
    ("Canadá", "Catar"): {"fecha": "18/06/2026", "hora": "15:00", "grupo": "B"},
    ("México", "República de Corea"): {"fecha": "18/06/2026", "hora": "19:00", "grupo": "A"},
    ("Brasil", "Haití"): {"fecha": "19/06/2026", "hora": "21:00", "grupo": "C"},
    ("Escocia", "Marruecos"): {"fecha": "19/06/2026", "hora": "18:00", "grupo": "C"},
    ("Turquía", "Paraguay"): {"fecha": "19/06/2026", "hora": "20:00", "grupo": "D"},
    ("Estados Unidos", "Australia"): {"fecha": "19/06/2026", "hora": "12:00", "grupo": "D"},
    ("Alemania", "Costa de Marfil"): {"fecha": "20/06/2026", "hora": "16:00", "grupo": "E"},
    ("Ecuador", "Curazao"): {"fecha": "20/06/2026", "hora": "19:00", "grupo": "E"},
    ("Países Bajos", "Suecia"): {"fecha": "20/06/2026", "hora": "12:00", "grupo": "F"},
    ("Túnez", "Japón"): {"fecha": "20/06/2026", "hora": "22:00", "grupo": "F"},
    ("Uruguay", "Cabo Verde"): {"fecha": "21/06/2026", "hora": "18:00", "grupo": "H"},
    ("España", "Arabia Saudí"): {"fecha": "21/06/2026", "hora": "12:00", "grupo": "H"},
    ("Bélgica", "Irán"): {"fecha": "21/06/2026", "hora": "12:00", "grupo": "G"},
    ("Nueva Zelanda", "Egipto"): {"fecha": "21/06/2026", "hora": "18:00", "grupo": "G"},
    ("Noruega", "Senegal"): {"fecha": "22/06/2026", "hora": "20:00", "grupo": "I"},
    ("Francia", "Irak"): {"fecha": "22/06/2026", "hora": "17:00", "grupo": "I"},
    ("Argentina", "Austria"): {"fecha": "22/06/2026", "hora": "12:00", "grupo": "J"},
    ("Jordania", "Argelia"): {"fecha": "22/06/2026", "hora": "20:00", "grupo": "J"},
    ("Inglaterra", "Ghana"): {"fecha": "23/06/2026", "hora": "16:00", "grupo": "L"},
    ("Panamá", "Croacia"): {"fecha": "23/06/2026", "hora": "19:00", "grupo": "L"},
    ("Portugal", "Uzbekistán"): {"fecha": "23/06/2026", "hora": "12:00", "grupo": "K"},
    ("Colombia", "RD Congo"): {"fecha": "23/06/2026", "hora": "20:00", "grupo": "K"},
    ("Escocia", "Brasil"): {"fecha": "24/06/2026", "hora": "18:00", "grupo": "C"},
    ("Marruecos", "Haití"): {"fecha": "24/06/2026", "hora": "18:00", "grupo": "C"},
    ("Suiza", "Canadá"): {"fecha": "24/06/2026", "hora": "12:00", "grupo": "B"},
    ("Bosnia y Herzegovina", "Catar"): {"fecha": "24/06/2026", "hora": "12:00", "grupo": "B"},
    ("República Checa", "México"): {"fecha": "24/06/2026", "hora": "19:00", "grupo": "A"},
    ("Sudáfrica", "República de Corea"): {"fecha": "24/06/2026", "hora": "19:00", "grupo": "A"},
    ("Curazao", "Costa de Marfil"): {"fecha": "25/06/2026", "hora": "16:00", "grupo": "E"},
    ("Ecuador", "Alemania"): {"fecha": "25/06/2026", "hora": "16:00", "grupo": "E"},
    ("Japón", "Suecia"): {"fecha": "25/06/2026", "hora": "18:00", "grupo": "F"},
    ("Túnez", "Países Bajos"): {"fecha": "25/06/2026", "hora": "18:00", "grupo": "F"},
    ("Turquía", "Estados Unidos"): {"fecha": "25/06/2026", "hora": "19:00", "grupo": "D"},
    ("Paraguay", "Australia"): {"fecha": "25/06/2026", "hora": "19:00", "grupo": "D"},
    ("Noruega", "Francia"): {"fecha": "26/06/2026", "hora": "15:00", "grupo": "I"},
    ("Senegal", "Irak"): {"fecha": "26/06/2026", "hora": "15:00", "grupo": "I"},
    ("Egipto", "Irán"): {"fecha": "26/06/2026", "hora": "20:00", "grupo": "G"},
    ("Nueva Zelanda", "Bélgica"): {"fecha": "26/06/2026", "hora": "20:00", "grupo": "G"},
    ("Cabo Verde", "Arabia Saudí"): {"fecha": "26/06/2026", "hora": "19:00", "grupo": "H"},
    ("Uruguay", "España"): {"fecha": "26/06/2026", "hora": "18:00", "grupo": "H"},
    ("Panamá", "Inglaterra"): {"fecha": "27/06/2026", "hora": "17:00", "grupo": "L"},
    ("Croacia", "Ghana"): {"fecha": "27/06/2026", "hora": "17:00", "grupo": "L"},
    ("Argelia", "Austria"): {"fecha": "27/06/2026", "hora": "21:00", "grupo": "J"},
    ("Jordania", "Argentina"): {"fecha": "27/06/2026", "hora": "21:00", "grupo": "J"},
    ("Colombia", "Portugal"): {"fecha": "27/06/2026", "hora": "19:30", "grupo": "K"},
    ("RD Congo", "Uzbekistán"): {"fecha": "27/06/2026", "hora": "19:30", "grupo": "K"},
}

KNOCKOUT_FIXTURES: list[dict] = [
    {"id": 73, "fecha": "28/06/2026", "hora": "12:00", "local": "Ganador 2A", "visitante": "Ganador 2B", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 74, "fecha": "29/06/2026", "hora": "16:30", "local": "Ganador 1E", "visitante": "Ganador 3ABCDF", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 75, "fecha": "29/06/2026", "hora": "19:00", "local": "Ganador 1F", "visitante": "Ganador 2C", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 76, "fecha": "29/06/2026", "hora": "12:00", "local": "Ganador 1C", "visitante": "Ganador 2F", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 77, "fecha": "30/06/2026", "hora": "17:00", "local": "Ganador 1I", "visitante": "Ganador 3CDFGH", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 78, "fecha": "30/06/2026", "hora": "12:00", "local": "Ganador 2E", "visitante": "Ganador 2I", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 79, "fecha": "30/06/2026", "hora": "19:00", "local": "Ganador 1A", "visitante": "Ganador 3CEFHI", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 80, "fecha": "01/07/2026", "hora": "12:00", "local": "Ganador 1L", "visitante": "Ganador 3EHIJK", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 81, "fecha": "01/07/2026", "hora": "17:00", "local": "Ganador 1D", "visitante": "Ganador 3BEFIJ", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 82, "fecha": "01/07/2026", "hora": "13:00", "local": "Ganador 1G", "visitante": "Ganador 3AEHIJ", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 83, "fecha": "02/07/2026", "hora": "19:00", "local": "Ganador 2K", "visitante": "Ganador 2L", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 84, "fecha": "02/07/2026", "hora": "12:00", "local": "Ganador 1H", "visitante": "Ganador 2J", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 85, "fecha": "02/07/2026", "hora": "20:00", "local": "Ganador 1B", "visitante": "Ganador 3EFGIJ", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 86, "fecha": "03/07/2026", "hora": "18:00", "local": "Ganador 1J", "visitante": "Ganador 2H", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 87, "fecha": "03/07/2026", "hora": "20:30", "local": "Ganador 1K", "visitante": "Ganador 3DEIJL", "fase": "Dieciseisavos", "jornada": 13},
    {"id": 88, "fecha": "03/07/2026", "hora": "13:00", "local": "Ganador 2D", "visitante": "Ganador 2G", "fase": "Dieciseisavos", "jornada": 13},
]

# Dieciseisavos (partidos 73–88): equipos reales post-fase de grupos (orden = IDs 73…88).
R32_RESOLVED: list[tuple[str, str]] = [
    ("Sudáfrica", "Canadá"),
    ("Alemania", "Paraguay"),
    ("Países Bajos", "Marruecos"),
    ("Brasil", "Japón"),
    ("Francia", "Suecia"),
    ("Costa de Marfil", "Noruega"),
    ("México", "Ecuador"),
    ("Inglaterra", "RD Congo"),
    ("Estados Unidos", "Bosnia y Herzegovina"),
    ("Bélgica", "Senegal"),
    ("Portugal", "Croacia"),
    ("España", "Austria"),
    ("Suiza", "Argelia"),
    ("Argentina", "Cabo Verde"),
    ("Colombia", "Ghana"),
    ("Australia", "Egipto"),
]

# Placeholders al estilo porra (Ganador 1A vs Ganador 2B) — octavos en adelante.

JORNADA_BY_GROUP = {letter: idx for idx, letter in enumerate("ABCDEFGHIJKL", start=1)}


def jornada_for_group(grupo: str) -> int:
    return JORNADA_BY_GROUP.get(grupo, 0)


def lookup_pair_meta(local: str, visitante: str) -> dict[str, str]:
    key = (local.strip(), visitante.strip())
    if key not in FIFA_PAIR_META:
        raise KeyError(f"Pareja no encontrada en calendario FIFA: {local} vs {visitante}")
    return FIFA_PAIR_META[key]


def build_all_matches(
    group_fixtures: list[tuple[str, str]] | None = None,
) -> list[dict]:
    """Genera los 104 partidos con metadata oficial."""
    fixtures = group_fixtures or PORRA_GROUP_FIXTURES
    matches: list[dict] = []

    for i, (local, visitante) in enumerate(fixtures, start=1):
        meta = lookup_pair_meta(local, visitante)
        grupo = meta["grupo"]
        matches.append(
            {
                "id": i,
                "fecha": meta["fecha"],
                "hora": meta["hora"],
                "local": local,
                "visitante": visitante,
                "grupo": grupo,
                "fase": "Grupos",
                "jornada": jornada_for_group(grupo),
            }
        )

    for idx, ko in enumerate(KNOCKOUT_FIXTURES):
        if idx < len(R32_RESOLVED):
            local, visitante = R32_RESOLVED[idx]
        else:
            raise ValueError(f"Falta equipo R32 para partido {ko['id']}")
        matches.append(
            {
                "id": ko["id"],
                "fecha": ko["fecha"],
                "hora": ko["hora"],
                "local": local,
                "visitante": visitante,
                "grupo": "",
                "fase": ko["fase"],
                "jornada": ko["jornada"],
            }
        )

    octavos_start = 89
    for i in range(8):
        mid = octavos_start + i
        ko_meta = _KO_REST[mid]
        matches.append(
            {
                "id": mid,
                "fecha": ko_meta["fecha"],
                "hora": ko_meta["hora"],
                "local": f"Ganador Octavos {i * 2 + 1}",
                "visitante": f"Ganador Octavos {i * 2 + 2}",
                "grupo": "",
                "fase": "Octavos",
                "jornada": 14,
            }
        )

    cuartos_start = 97
    for i in range(4):
        mid = cuartos_start + i
        ko_meta = _KO_REST[mid]
        matches.append(
            {
                "id": mid,
                "fecha": ko_meta["fecha"],
                "hora": ko_meta["hora"],
                "local": f"Ganador Cuartos {i * 2 + 1}",
                "visitante": f"Ganador Cuartos {i * 2 + 2}",
                "grupo": "",
                "fase": "Cuartos",
                "jornada": 15,
            }
        )

    for i in range(2):
        mid = 101 + i
        ko_meta = _KO_REST[mid]
        matches.append(
            {
                "id": mid,
                "fecha": ko_meta["fecha"],
                "hora": ko_meta["hora"],
                "local": f"Ganador Semifinal {i * 2 + 1}",
                "visitante": f"Ganador Semifinal {i * 2 + 2}",
                "grupo": "",
                "fase": "Semifinal",
                "jornada": 16,
            }
        )

    ko_meta = _KO_REST[103]
    matches.append(
        {
            "id": 103,
            "fecha": ko_meta["fecha"],
            "hora": ko_meta["hora"],
            "local": "Perdedor Semifinal 1",
            "visitante": "Perdedor Semifinal 2",
            "grupo": "",
            "fase": "Tercer puesto",
            "jornada": 17,
        }
    )

    ko_meta = _KO_REST[104]
    matches.append(
        {
            "id": 104,
            "fecha": ko_meta["fecha"],
            "hora": ko_meta["hora"],
            "local": "Ganador Semifinal 1",
            "visitante": "Ganador Semifinal 2",
            "grupo": "",
            "fase": "Final",
            "jornada": 17,
        }
    )

    return matches


_KO_REST: dict[int, dict[str, str]] = {
    89: {"fecha": "04/07/2026", "hora": "17:00"},
    90: {"fecha": "04/07/2026", "hora": "12:00"},
    91: {"fecha": "05/07/2026", "hora": "16:00"},
    92: {"fecha": "05/07/2026", "hora": "18:00"},
    93: {"fecha": "06/07/2026", "hora": "14:00"},
    94: {"fecha": "06/07/2026", "hora": "17:00"},
    95: {"fecha": "07/07/2026", "hora": "12:00"},
    96: {"fecha": "07/07/2026", "hora": "13:00"},
    97: {"fecha": "09/07/2026", "hora": "16:00"},
    98: {"fecha": "10/07/2026", "hora": "12:00"},
    99: {"fecha": "11/07/2026", "hora": "17:00"},
    100: {"fecha": "11/07/2026", "hora": "20:00"},
    101: {"fecha": "14/07/2026", "hora": "14:00"},
    102: {"fecha": "15/07/2026", "hora": "15:00"},
    103: {"fecha": "18/07/2026", "hora": "17:00"},
    104: {"fecha": "19/07/2026", "hora": "15:00"},
}
