#!/usr/bin/env python3
"""Genera cromos tipo Panini para la porra del Mundial 2026."""

from __future__ import annotations

import argparse
import base64
import json
import os
import sys
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config" / "cromos.json"
PROMPT_PATH = ROOT / "prompts" / "cromo_portrait.txt"
OUTPUT_DIR = ROOT / "output" / "cromos"
PORTRAIT_DIR = OUTPUT_DIR / "_portraits"

CARD_W, CARD_H = 750, 1050
INNER_MARGIN = 18
BORDER = 14

RED_ES = "#C60B1E"
YELLOW_ES = "#FFC400"
GOLD = "#D4AF37"
NAVY = "#1B2A4A"
WHITE = "#FFFFFF"


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def get_player(config: dict, player_id: str) -> dict:
    for player in config["players"]:
        if player["id"] == player_id:
            return player
    ids = ", ".join(p["id"] for p in config["players"])
    raise SystemExit(f"Jugador '{player_id}' no encontrado. Disponibles: {ids}")


def load_prompt() -> str:
    return PROMPT_PATH.read_text(encoding="utf-8").strip()


def portrait_path(player_id: str) -> Path:
    return PORTRAIT_DIR / f"{player_id}_portrait.png"


def cromo_path(player_id: str) -> Path:
    return OUTPUT_DIR / f"{player_id}.png"


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates = [
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/segoeuib.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        ]
    else:
        candidates = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
        ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def _rounded_rect_mask(size: tuple[int, int], radius: int) -> Image.Image:
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius=radius, fill=255)
    return mask


def _draw_spain_flag(draw: ImageDraw.ImageDraw, x: int, y: int, w: int, h: int) -> None:
    stripe_h = h // 3
    draw.rectangle((x, y, x + w, y + stripe_h), fill=RED_ES)
    draw.rectangle((x, y + stripe_h, x + w, y + 2 * stripe_h), fill=YELLOW_ES)
    draw.rectangle((x, y + 2 * stripe_h, x + w, y + h), fill=RED_ES)
    draw.rectangle((x, y, x + w, y + h), outline=GOLD, width=2)


def _draw_rfef_crest(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int) -> None:
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=RED_ES, outline=GOLD, width=3)
    draw.ellipse((cx - r + 8, cy - r + 8, cx + r - 8, cy + r - 8), outline=YELLOW_ES, width=2)
    font = _font(max(12, r // 2), bold=True)
    draw.text((cx, cy), "RFEF", fill=WHITE, font=font, anchor="mm")


def _load_stadium_background() -> Image.Image:
    for rel in (
        "assets/wc2026_stadium.jpg",
        "assets/wc2026_hero.jpg",
        "assets/stadium.jpg",
        "assets/wc2026_crowd.jpg",
        "assets/wc2026_hosts.jpg",
        "assets/football.jpg",
    ):
        path = ROOT / rel
        if not path.exists() or path.stat().st_size < 1000:
            continue
        try:
            bg = Image.open(path).convert("RGB")
        except Exception:
            continue
        bg = ImageOps.fit(bg, (CARD_W, CARD_H), method=Image.Resampling.LANCZOS)
        bg = bg.filter(ImageFilter.GaussianBlur(radius=6))
        overlay = Image.new("RGBA", (CARD_W, CARD_H))
        ov_draw = ImageDraw.Draw(overlay)
        ov_draw.rectangle((0, 0, CARD_W, CARD_H // 2), fill=(198, 11, 30, 140))
        ov_draw.rectangle((0, CARD_H // 2, CARD_W, CARD_H), fill=(27, 42, 74, 170))
        bg = Image.alpha_composite(bg.convert("RGBA"), overlay).convert("RGB")
        return bg
    grad = Image.new("RGB", (CARD_W, CARD_H), RED_ES)
    draw = ImageDraw.Draw(grad)
    for y in range(CARD_H):
        t = y / CARD_H
        r = int(198 * (1 - t) + 27 * t)
        g = int(11 * (1 - t) + 42 * t)
        b = int(30 * (1 - t) + 74 * t)
        draw.line((0, y, CARD_W, y), fill=(r, g, b))
    return grad


def build_card_frame() -> Image.Image:
    """Marco Panini reutilizable (fondo + borde, sin retrato ni textos de jugador)."""
    card = _load_stadium_background().convert("RGBA")
    draw = ImageDraw.Draw(card)

    inner = (
        INNER_MARGIN + BORDER,
        INNER_MARGIN + BORDER,
        CARD_W - INNER_MARGIN - BORDER,
        CARD_H - INNER_MARGIN - BORDER,
    )
    draw.rounded_rectangle(inner, radius=22, outline=GOLD, width=4)
    draw.rounded_rectangle(
        (inner[0] + 6, inner[1] + 6, inner[2] - 6, inner[3] - 6),
        radius=18,
        outline=RED_ES,
        width=2,
    )

    for x, y in (
        (INNER_MARGIN, INNER_MARGIN),
        (CARD_W - INNER_MARGIN, INNER_MARGIN),
        (INNER_MARGIN, CARD_H - INNER_MARGIN),
        (CARD_W - INNER_MARGIN, CARD_H - INNER_MARGIN),
    ):
        draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill=GOLD, outline=WHITE, width=1)

    return card


def _fit_portrait(portrait: Image.Image) -> Image.Image:
    """Recorta y escala el retrato para la zona superior del cromo."""
    target_w = CARD_W - 2 * (INNER_MARGIN + BORDER + 20)
    target_h = int(CARD_H * 0.58)
    portrait = portrait.convert("RGBA")

    ratio = max(target_w / portrait.width, target_h / portrait.height)
    new_size = (int(portrait.width * ratio), int(portrait.height * ratio))
    portrait = portrait.resize(new_size, Image.Resampling.LANCZOS)

    left = (portrait.width - target_w) // 2
    top = max(0, portrait.height - target_h - int(portrait.height * 0.05))
    portrait = portrait.crop((left, top, left + target_w, top + target_h))

    fade = Image.new("L", portrait.size, 255)
    fade_draw = ImageDraw.Draw(fade)
    for y in range(portrait.height):
        if y > portrait.height - 80:
            alpha = int(255 * max(0, (portrait.height - y) / 80))
            fade_draw.line((0, y, portrait.width, y), fill=alpha)
    portrait.putalpha(fade)
    return portrait


def compose_cromo(portrait: Image.Image, metadata: dict) -> Image.Image:
    """Compone el cromo final sobre el marco fijo."""
    card = build_card_frame()
    draw = ImageDraw.Draw(card)

    fitted = _fit_portrait(portrait)
    px = (CARD_W - fitted.width) // 2
    py = INNER_MARGIN + BORDER + 50
    card.paste(fitted, (px, py), fitted)

    _draw_spain_flag(draw, INNER_MARGIN + BORDER + 16, INNER_MARGIN + BORDER + 16, 72, 48)

    title_font = _font(22, bold=True)
    draw.text(
        (CARD_W - INNER_MARGIN - BORDER - 20, INNER_MARGIN + BORDER + 24),
        "MUNDIAL",
        fill=WHITE,
        font=title_font,
        anchor="ra",
    )
    draw.text(
        (CARD_W - INNER_MARGIN - BORDER - 20, INNER_MARGIN + BORDER + 50),
        "2026",
        fill=YELLOW_ES,
        font=_font(28, bold=True),
        anchor="ra",
    )

    band_top = CARD_H - 200
    draw.rectangle(
        (INNER_MARGIN + BORDER, band_top, CARD_W - INNER_MARGIN - BORDER, CARD_H - INNER_MARGIN - BORDER),
        fill=RED_ES,
    )
    draw.rectangle(
        (INNER_MARGIN + BORDER, band_top, CARD_W - INNER_MARGIN - BORDER, band_top + 6),
        fill=YELLOW_ES,
    )

    name = metadata["name"].upper()
    name_font = _font(52, bold=True)
    draw.text((CARD_W // 2, band_top + 48), name, fill=WHITE, font=name_font, anchor="mm")

    position = metadata["position"]
    pos_font = _font(20, bold=True)
    badge_w = len(position) * 14 + 40
    badge_x = CARD_W // 2 - badge_w // 2
    badge_y = band_top + 88
    draw.rounded_rectangle(
        (badge_x, badge_y, badge_x + badge_w, badge_y + 34),
        radius=17,
        fill=YELLOW_ES,
        outline=GOLD,
        width=2,
    )
    draw.text((CARD_W // 2, badge_y + 17), position, fill=NAVY, font=pos_font, anchor="mm")

    number = metadata.get("number", 0)
    num_text = f"#{number:02d}"
    draw.text(
        (INNER_MARGIN + BORDER + 24, CARD_H - INNER_MARGIN - BORDER - 24),
        num_text,
        fill=GOLD,
        font=_font(26, bold=True),
        anchor="ls",
    )

    _draw_rfef_crest(draw, CARD_W - INNER_MARGIN - BORDER - 36, CARD_H - INNER_MARGIN - BORDER - 36, 28)

    draw.rounded_rectangle(
        (INNER_MARGIN, INNER_MARGIN, CARD_W - INNER_MARGIN, CARD_H - INNER_MARGIN),
        radius=26,
        outline=GOLD,
        width=BORDER,
    )

    mask = _rounded_rect_mask((CARD_W, CARD_H), 26)
    card.putalpha(mask)
    return card


def _create_fallback_portrait(photo_path: Path) -> Image.Image:
    """Retrato estilizado sin IA: camiseta España simplificada sobre la foto original."""
    photo = Image.open(photo_path).convert("RGBA")
    size = 1024
    photo = ImageOps.fit(photo, (size, size), method=Image.Resampling.LANCZOS)

    bg = Image.new("RGBA", (size, size), (30, 20, 50, 255))
    stadium = _load_stadium_background().resize((size, size), Image.Resampling.LANCZOS)
    bg = Image.blend(stadium.convert("RGBA"), bg, alpha=0.35)

    subject = photo.copy()
    w, h = subject.size
    jersey_top = int(h * 0.52)
    overlay = Image.new("RGBA", subject.size, (0, 0, 0, 0))
    ov = ImageDraw.Draw(overlay)
    ov.rectangle((int(w * 0.08), jersey_top, int(w * 0.92), h), fill=(198, 11, 30, 225))
    ov.polygon(
        [
            (int(w * 0.28), jersey_top),
            (int(w * 0.72), jersey_top),
            (int(w * 0.68), jersey_top - 22),
            (int(w * 0.32), jersey_top - 22),
        ],
        fill=(255, 196, 0, 240),
    )
    ov.ellipse(
        (w // 2 - 38, jersey_top + 36, w // 2 + 38, jersey_top + 112),
        fill=(140, 8, 20, 250),
        outline=(255, 196, 0, 255),
        width=3,
    )
    ov.text((w // 2, jersey_top + 74), "RFEF", fill=(255, 255, 255, 255), font=_font(20, bold=True), anchor="mm")
    subject = Image.alpha_composite(subject, overlay)

    result = Image.alpha_composite(bg, subject)
    return result.convert("RGBA")


def generate_portrait(photo_path: Path, prompt: str, *, allow_fallback: bool = True) -> Image.Image:
    """Genera retrato con equipación España vía OpenAI Images API."""
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        if allow_fallback:
            print("OPENAI_API_KEY no configurada — usando retrato estilizado (fallback).")
            print("Para equipación realista, configura la API y vuelve a ejecutar.")
            return _create_fallback_portrait(photo_path)
        raise SystemExit(
            "OPENAI_API_KEY no configurada.\n"
            "Opciones:\n"
            "  1. Configura la variable y vuelve a ejecutar.\n"
            "  2. Genera el retrato manualmente con prompts/cromo_portrait.txt\n"
            "     y guárdalo en output/cromos/_portraits/<id>_portrait.png\n"
            "  3. Usa --compose-only con un retrato ya existente."
        )

    try:
        from openai import OpenAI
    except ImportError as exc:
        raise SystemExit("Instala openai: py -m pip install openai") from exc

    client = OpenAI(api_key=api_key)
    photo_bytes = photo_path.read_bytes()

    models = ("gpt-image-1", "dall-e-2")
    last_error: Exception | None = None

    for model in models:
        try:
            kwargs: dict = {
                "model": model,
                "image": photo_bytes,
                "prompt": prompt,
                "n": 1,
            }
            if model == "dall-e-2":
                kwargs["size"] = "1024x1024"
            response = client.images.edit(**kwargs)
            item = response.data[0]
            if item.b64_json:
                raw = base64.b64decode(item.b64_json)
            elif item.url:
                import urllib.request

                with urllib.request.urlopen(item.url) as resp:
                    raw = resp.read()
            else:
                raise RuntimeError("La API no devolvió imagen")
            return Image.open(BytesIO(raw)).convert("RGBA")
        except Exception as exc:  # noqa: BLE001 — probar modelos alternativos
            last_error = exc
            continue

    raise SystemExit(f"Error al generar retrato con OpenAI: {last_error}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genera cromos Panini Mundial 2026")
    parser.add_argument("--player", required=True, help="ID del jugador (ej. felipe)")
    parser.add_argument(
        "--compose-only",
        action="store_true",
        help="Solo compone el marco sobre un retrato existente",
    )
    parser.add_argument(
        "--portrait-only",
        action="store_true",
        help="Solo genera el retrato IA, sin montar el cromo",
    )
    parser.add_argument(
        "--force-portrait",
        action="store_true",
        help="Regenera el retrato aunque ya exista",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = load_config()
    player = get_player(config, args.player)
    photo_path = ROOT / player["photo"]
    if not photo_path.exists():
        raise SystemExit(f"Foto no encontrada: {photo_path}")

    prompt = load_prompt()
    p_path = portrait_path(player["id"])

    if args.compose_only:
        if not p_path.exists():
            raise SystemExit(
                f"Retrato no encontrado: {p_path}\n"
                "Genera uno con IA o ejecuta sin --compose-only."
            )
        portrait = Image.open(p_path)
    elif p_path.exists() and not args.portrait_only and not args.force_portrait:
        print(f"Usando retrato existente: {p_path}")
        portrait = Image.open(p_path)
    else:
        print(f"Generando retrato IA para {player['name']}...")
        portrait = generate_portrait(photo_path, prompt)
        PORTRAIT_DIR.mkdir(parents=True, exist_ok=True)
        portrait.save(p_path, "PNG")
        print(f"Retrato guardado: {p_path}")

    if args.portrait_only:
        return

    print(f"Componiendo cromo de {player['name']}...")
    cromo = compose_cromo(portrait, player)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    cromo.save(cromo_path(player["id"]), "PNG")
    print(f"Cromo: {cromo_path(player['id'])}")


if __name__ == "__main__":
    main()
