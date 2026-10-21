"""Build the footer strip images: the orange A mark and the A / E / S letters.

Outlook's Word engine pads text rows to the font's line spacing, so the
letters are images cropped to their ink; each row is then exactly the letter
height plus the cell padding. Images are 2x for high-DPI screens; the HTML
sets the display size.

Run from the install root:  python scripts/build_footer_strip_images.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

from aes.geolocate_headers import (  # noqa: E402
    AES_STRIP_AES_COLOR,
    AES_STRIP_BG,
    AES_STRIP_LETTER_H,
    AES_STRIP_MARK_H,
    AES_STRIP_MARK_W,
)

SCALE = 2
ICONS = ROOT / "assets" / "icons"


def _rgb(hex_colour: str) -> tuple[int, int, int]:
    h = hex_colour.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _bold_font(size: int) -> ImageFont.FreeTypeFont:
    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"
        / "Montserrat-VariableFont_wght.ttf",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "Montserrat-VariableFont_wght.ttf",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "Montserrat-Bold.ttf",
    ]
    for path in candidates:
        if path.is_file():
            font = ImageFont.truetype(str(path), size)
            try:
                font.set_variation_by_name("Bold")
            except Exception:
                pass
            return font
    print("Montserrat not found; using Segoe UI Bold", file=sys.stderr)
    return ImageFont.truetype(
        str(Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "segoeuib.ttf"), size
    )


def build_letter(letter: str) -> Path:
    bg, fg = _rgb(AES_STRIP_BG), _rgb(AES_STRIP_AES_COLOR)
    font = _bold_font(400)
    canvas = Image.new("RGB", (600, 600), bg)
    ImageDraw.Draw(canvas).text((100, 50), letter, font=font, fill=fg)
    mask = Image.new("L", canvas.size, 0)
    ImageDraw.Draw(mask).text((100, 50), letter, font=font, fill=255)
    box = mask.getbbox()
    if box is None:
        raise RuntimeError(f"no ink for {letter!r}")
    ink = canvas.crop(box)
    out_h = AES_STRIP_LETTER_H * SCALE
    out_w = max(1, round(ink.width * out_h / ink.height))
    out = ink.resize((out_w, out_h), Image.Resampling.LANCZOS)
    path = ICONS / f"aes_mark_letter_{letter.lower()}.png"
    out.save(path, optimize=True)
    return path


def build_mark() -> Path:
    src = Image.open(ROOT / "assets" / "brand" / "abyitself.ico")
    src.size = max(src.info["sizes"])  # type: ignore[misc]
    rgba = src.convert("RGBA")
    flat = Image.new("RGBA", rgba.size, _rgb(AES_STRIP_BG) + (255,))
    flat.alpha_composite(rgba)
    size = (AES_STRIP_MARK_W * SCALE, AES_STRIP_MARK_H * SCALE)
    out = flat.convert("RGB").resize(size, Image.Resampling.LANCZOS)
    path = ICONS / "aes_mark_footer.png"
    out.save(path, optimize=True)
    return path


def main() -> None:
    print(build_mark())
    for letter in "AES":
        path = build_letter(letter)
        w, h = Image.open(path).size
        print(path, f"{w // SCALE}x{h // SCALE} display")


if __name__ == "__main__":
    main()
