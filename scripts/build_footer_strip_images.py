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
    AES_STRIP_AES_COL,
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
    glyph = ink.resize((out_w, out_h), Image.Resampling.LANCZOS)
    # Every letter is the column width with its ink centred, so A, E and S
    # line up on one centre line.
    col_w = max(out_w, AES_STRIP_AES_COL * SCALE)
    out = Image.new("RGB", (col_w, out_h), bg)
    out.paste(glyph, ((col_w - out_w) // 2, 0))
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


def build_aura_chip() -> Path:
    """Inside of the 78x16 AURA chip: rainbow-gradient AURA, no frame.

    The HTML cell draws the 1px border, so the image is the 76x14 interior.
    """
    from aes.geolocate_headers import AES_CHIP_BG

    w, h = 76 * SCALE, 14 * SCALE
    bg = _rgb(AES_CHIP_BG)
    stops = (
        (0.00, (228, 3, 3)),
        (0.20, (255, 140, 0)),
        (0.40, (255, 237, 0)),
        (0.60, (0, 128, 38)),
        (0.80, (36, 64, 142)),
        (1.00, (115, 41, 130)),
    )

    def color_at(t: float) -> tuple[int, int, int]:
        t = max(0.0, min(1.0, t))
        for i in range(len(stops) - 1):
            t0, c0 = stops[i]
            t1, c1 = stops[i + 1]
            if t0 <= t <= t1:
                u = 0.0 if t1 == t0 else (t - t0) / (t1 - t0)
                return tuple(int(c0[j] + (c1[j] - c0[j]) * u) for j in range(3))  # type: ignore[return-value]
        return stops[-1][1]

    # Cap height ~80% of the interior; letters spread across ~80% of the width.
    text = "AURA"
    target_h = round(h * 0.8)
    size = target_h
    font = _bold_font(size)
    probe = ImageDraw.Draw(Image.new("L", (1, 1)))
    for _ in range(40):
        box = probe.textbbox((0, 0), "A", font=font)
        if box[3] - box[1] >= target_h:
            break
        size += 1
        font = _bold_font(size)
    glyphs = [probe.textbbox((0, 0), ch, font=font) for ch in text]
    widths = [b[2] - b[0] for b in glyphs]
    span = round(w * 0.8)
    gap = max(0, (span - sum(widths)) // (len(text) - 1))
    total = sum(widths) + gap * (len(text) - 1)
    mask = Image.new("L", (w, h), 0)
    draw = ImageDraw.Draw(mask)
    x0 = x = (w - total) // 2
    for ch, box, cw in zip(text, glyphs, widths):
        y = (h - (box[3] - box[1])) // 2 - box[1]
        draw.text((x - box[0], y), ch, font=font, fill=255)
        x += cw + gap
    # Gradient spans the lettering, not the chip, so it runs red to violet.
    grad = Image.new("RGB", (w, h), bg)
    px = grad.load()
    for ix in range(w):
        c = color_at((ix - x0) / max(1, total - 1))
        for iy in range(h):
            px[ix, iy] = c
    out = Image.new("RGB", (w, h), bg)
    out.paste(grad, (0, 0), mask)
    path = ICONS / "aes_aura_chip.png"
    out.save(path, optimize=True)
    return path


def main() -> None:
    print(build_mark())
    for letter in "AES":
        path = build_letter(letter)
        w, h = Image.open(path).size
        print(path, f"{w // SCALE}x{h // SCALE} display")
    aura = build_aura_chip()
    print(aura, "76x14 display")


if __name__ == "__main__":
    main()
