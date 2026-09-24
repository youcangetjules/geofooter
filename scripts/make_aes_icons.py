"""Generate the AES ribbon/toolbar icon set.

Outputs, for each icon name:
  <name>.png          32x32 RGBA (Office ribbon, real alpha glow)
  <name>.bmp          32x32 24-bit over white (legacy VBA CommandBar)
  <name>_mask.bmp     32x32 mask (white = transparent, black = opaque)
  <name>_16*.bmp      16px variants of the BMP pair

Everything is drawn at 8x supersampling and downscaled with Lanczos so the
"1 px" strokes stay crisp and the outer glow is smooth.

Copies the results into every folder AES resolves icons from.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

S = 8  # supersample factor
SIZE = 32
CANVAS = SIZE * S

GREEN = (34, 197, 94, 255)
GREEN_DARK = (22, 163, 74, 255)
RED = (239, 68, 68, 255)
AMBER = (245, 158, 11, 255)
ORANGE = (249, 115, 22, 255)
BLUE = (37, 99, 235, 255)
STEEL = (100, 116, 139, 255)
PAPER_EDGE = (148, 163, 184, 255)


def _blank(px: int = CANVAS) -> Image.Image:
    return Image.new("RGBA", (px, px), (0, 0, 0, 0))


def _with_glow(shape: Image.Image, color, radius_px: float = 2.4,
               strength: int = 110) -> Image.Image:
    """Put a soft outer glow (same hue) underneath the shape layer."""
    alpha = shape.split()[3]
    glow_a = alpha.filter(ImageFilter.GaussianBlur(radius_px * S))
    glow_a = glow_a.point(lambda a: min(255, int(a * strength / 100)))
    glow = Image.new("RGBA", shape.size, tuple(color[:3]) + (0,))
    glow.putalpha(glow_a)
    out = _blank(shape.size[0])
    out.alpha_composite(glow)
    out.alpha_composite(shape)
    return out


def _downscale(img: Image.Image, px: int = SIZE) -> Image.Image:
    return img.resize((px, px), Image.LANCZOS)


def _pt(x: float, y: float) -> tuple:
    return (x * S, y * S)


# ---------------------------------------------------------------- shapes

def icon_service_on() -> Image.Image:
    """Hollow green forward-pointing triangle, 1px stroke, outer glow."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    pts = [_pt(10.5, 6.5), _pt(10.5, 25.5), _pt(26.0, 16.0)]
    d.polygon(pts, outline=GREEN, width=S)  # 1px at final size
    return _with_glow(shape, GREEN)


def icon_service_off() -> Image.Image:
    """Hollow red square, 1px stroke, outer glow."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.rectangle([_pt(8.5, 8.5), _pt(23.5, 23.5)], outline=RED, width=S)
    return _with_glow(shape, RED)


def icon_service_busy() -> Image.Image:
    """Hollow amber circle (also 'ON but 0 inboxes' warning), matching style."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.ellipse([_pt(8.0, 8.0), _pt(24.0, 24.0)], outline=AMBER, width=S)
    return _with_glow(shape, AMBER)


def icon_short_scan() -> Image.Image:
    """Single orange bar — the one-line sibling of Full Scan."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.rounded_rectangle([_pt(4.0, 13.5), _pt(28.0, 18.5)],
                        radius=2.5 * S, fill=ORANGE)
    return _with_glow(shape, ORANGE, radius_px=1.2, strength=50)


def _bars(d: ImageDraw.ImageDraw) -> None:
    """Three big rounded orange bars (Full Scan motif, enlarged)."""
    for cy in (7.0, 16.0, 25.0):
        d.rounded_rectangle(
            [_pt(4.0, cy - 2.5), _pt(28.0, cy + 2.5)],
            radius=2.5 * S, fill=ORANGE,
        )


def icon_full_scan() -> Image.Image:
    shape = _blank()
    _bars(ImageDraw.Draw(shape))
    return _with_glow(shape, ORANGE, radius_px=1.2, strength=50)


def _plus_badge(shape: Image.Image, d: ImageDraw.ImageDraw) -> None:
    """Blue circle badge with a white '+', bottom-right, hole-punched."""
    hole = _blank()
    hd = ImageDraw.Draw(hole)
    hd.ellipse([_pt(13.0, 13.0), _pt(31.5, 31.5)], fill=(0, 0, 0, 255))
    shape.paste((0, 0, 0, 0), (0, 0), hole)
    d.ellipse([_pt(14.5, 14.5), _pt(30.0, 30.0)], fill=BLUE)
    cx, cy, r = 22.25, 22.25, 4.2
    w = int(1.9 * S)
    d.line([_pt(cx - r, cy), _pt(cx + r, cy)], fill=(255, 255, 255, 255), width=w)
    d.line([_pt(cx, cy - r), _pt(cx, cy + r)], fill=(255, 255, 255, 255), width=w)


def icon_deep_scan() -> Image.Image:
    """Full-scan bars plus a blue circle badge with a '+' inside."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    _bars(d)
    _plus_badge(shape, d)
    return _with_glow(shape, ORANGE, radius_px=1.2, strength=45)


def icon_diagnostics() -> Image.Image:
    """Steel spanner/wrench, drawn for a 16px normal-size ribbon button."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    # handle: thick rounded diagonal
    d.line([_pt(10.5, 10.5), _pt(25.0, 25.0)], fill=STEEL, width=int(5.5 * S))
    d.ellipse([_pt(22.4, 22.4), _pt(27.6, 27.6)], fill=STEEL)
    # head: solid disc at top-left …
    d.ellipse([_pt(2.0, 2.0), _pt(19.0, 19.0)], fill=STEEL)
    # … carved into an open C-jaw: cut the bolt slot toward the top-left
    cut = _blank()
    cd = ImageDraw.Draw(cut)
    cd.polygon([_pt(0.0, 0.0), _pt(13.0, 0.0), _pt(8.5, 8.0),
                _pt(9.0, 12.0), _pt(0.0, 13.0)], fill=(0, 0, 0, 255))
    cd.ellipse([_pt(5.2, 5.2), _pt(13.2, 13.2)], fill=(0, 0, 0, 255))
    shape.paste((0, 0, 0, 0), (0, 0), cut)
    return _with_glow(shape, STEEL, radius_px=0.8, strength=35)


def icon_view_logs() -> Image.Image:
    """Log-file page: bordered sheet with timestamped-looking lines (16px)."""
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.rounded_rectangle([_pt(5.0, 2.0), _pt(27.0, 30.0)], radius=2 * S,
                        fill=(255, 255, 255, 255), outline=PAPER_EDGE, width=S)
    rows = [(7.0, GREEN_DARK), (12.5, STEEL), (18.0, STEEL), (23.5, AMBER)]
    for y, color in rows:
        d.rounded_rectangle([_pt(8.0, y), _pt(12.5, y + 2.6)],
                            radius=1.0 * S, fill=color)
        d.rounded_rectangle([_pt(14.5, y), _pt(24.0, y + 2.6)],
                            radius=1.0 * S, fill=(203, 213, 225, 255))
    return _with_glow(shape, STEEL, radius_px=0.6, strength=25)


def icon_guri() -> Image.Image:
    """GURI — teal rounded tile with a bold G (brand accent)."""
    TEAL = (15, 107, 124, 255)  # #0f6b7c
    TEAL_SOFT = (227, 241, 244, 255)
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.rounded_rectangle(
        [_pt(4.0, 4.0), _pt(28.0, 28.0)],
        radius=5 * S,
        fill=TEAL_SOFT,
        outline=TEAL,
        width=max(1, S),
    )
    # Stylized G from arcs + bar (readable at 32px).
    d.arc([_pt(9.0, 8.0), _pt(23.0, 24.0)], start=40, end=320, fill=TEAL, width=2 * S)
    d.line([_pt(16.0, 16.0), _pt(23.0, 16.0)], fill=TEAL, width=2 * S)
    d.line([_pt(23.0, 16.0), _pt(23.0, 20.5)], fill=TEAL, width=2 * S)
    return _with_glow(shape, TEAL, radius_px=1.4, strength=70)


def icon_aura() -> Image.Image:
    """Aura — privacy aura rings + scrub stroke (data-broker removal)."""
    VIOLET = (109, 40, 217, 255)
    VIOLET_SOFT = (167, 139, 250, 255)
    shape = _blank()
    d = ImageDraw.Draw(shape)
    d.ellipse([_pt(4.5, 4.5), _pt(27.5, 27.5)], outline=VIOLET_SOFT, width=max(1, S))
    d.ellipse([_pt(8.0, 8.0), _pt(24.0, 24.0)], outline=VIOLET, width=max(1, int(1.25 * S)))
    d.line([_pt(11.0, 21.0), _pt(21.0, 11.0)], fill=VIOLET, width=2 * S)
    d.line([_pt(12.5, 22.5), _pt(14.5, 20.5)], fill=VIOLET_SOFT, width=max(1, S))
    return _with_glow(shape, VIOLET, radius_px=1.6, strength=75)


# ---------------------------------------------------------------- output

def bmp_pair(png: Image.Image, px: int):
    """(image over white, mask) as 24-bit-ready RGB images at px size."""
    img = png if png.size[0] == px else png.resize((px, px), Image.LANCZOS)
    over_white = Image.new("RGB", (px, px), (255, 255, 255))
    over_white.paste(img, (0, 0), img)
    alpha = img.split()[3]
    # mask: white where transparent, black where the icon is drawn
    mask = alpha.point(lambda a: 0 if a >= 96 else 255).convert("RGB")
    return over_white, mask


# name -> (draw fn returning the 256px hi-res image, ribbon button size)
ICONS = {
    "aes_service_on": (icon_service_on, 32),
    "aes_service_off": (icon_service_off, 32),
    "aes_service_busy": (icon_service_busy, 32),
    "aes_short_scan": (icon_short_scan, 32),
    "aes_full_scan": (icon_full_scan, 32),
    "aes_deep_scan": (icon_deep_scan, 32),
    "aes_guri": (icon_guri, 32),
    "aes_aura": (icon_aura, 32),
    "aes_diagnostics": (icon_diagnostics, 16),
    "aes_view_logs": (icon_view_logs, 16),
}

TARGET_DIRS = None  # filled in main() from geofooter_paths


def main() -> None:
    # Allow `python scripts/make_aes_icons.py` to import suite modules.
    root_guess = Path(__file__).resolve().parent.parent
    if str(root_guess) not in sys.path:
        sys.path.insert(0, str(root_guess))
    try:
        from geofooter_paths import get_install_root, icons_dir, vba_dir

        root = get_install_root()
        out = root / "assets" / "icons_build"
        targets = [
            Path.home() / "AppData/Local/GeoFooter/icons",
            vba_dir() / "icons",
            icons_dir(),
        ]
    except Exception:
        root = root_guess
        out = root / "assets" / "icons_build"
        targets = [
            Path.home() / "AppData/Local/GeoFooter/icons",
            root / "VBA" / "icons",
            root / "assets" / "icons",
        ]

    out.mkdir(parents=True, exist_ok=True)

    for name, (fn, base) in ICONS.items():
        hires = fn()
        # PNG at 2x the nominal button size: Office scales DOWN on high-DPI
        # ribbons instead of up, which keeps thin 1px strokes round and crisp.
        _downscale(hires, base * 2).save(out / f"{name}.png")
        bmp, mask = bmp_pair(hires, base)
        bmp.save(out / f"{name}.bmp")
        mask.save(out / f"{name}_mask.bmp")
        bmp16, mask16 = bmp_pair(hires, 16)
        bmp16.save(out / f"{name}_16.bmp")
        mask16.save(out / f"{name}_16_mask.bmp")

    # legacy fallback name used by the ribbon host / toolbar
    for suffix in (".png", ".bmp", "_mask.bmp", "_16.bmp", "_16_mask.bmp"):
        src = out / f"aes_short_scan{suffix}"
        if src.exists():
            shutil.copy2(src, out / f"aes_short_scan_drawn{suffix}")

    for target in targets:
        target.mkdir(parents=True, exist_ok=True)
        for f in out.iterdir():
            shutil.copy2(f, target / f.name)
        print(f"copied {len(list(out.iterdir()))} files -> {target}")


if __name__ == "__main__":
    main()
