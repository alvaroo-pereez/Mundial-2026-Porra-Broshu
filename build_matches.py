"""Generate data/matches_2026.json with 104 World Cup 2026 fixture slots."""
import json
from pathlib import Path

from fifa_schedule_2026 import build_all_matches


def main() -> None:
    matches = build_all_matches()
    out = Path(__file__).parent / "data" / "matches_2026.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)
    print(f"Written {len(matches)} matches to {out}")


if __name__ == "__main__":
    main()
