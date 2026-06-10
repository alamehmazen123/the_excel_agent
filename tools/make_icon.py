"""Generate ui/resources/app.ico -- an Excel-AI themed app icon.

Design: a rounded green tile (Excel green) with a white spreadsheet grid and a
small AI 'spark/sparkle' in the corner, evoking 'intelligent spreadsheet agent'.
Rendered large and downscaled into a multi-resolution .ico.
"""
from __future__ import annotations

import os

from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "ui", "resources", "app.ico")

EXCEL_GREEN = (33, 115, 70)      # deep Excel green
EXCEL_GREEN2 = (16, 124, 65)
GRID = (255, 255, 255)
SPARK = (255, 214, 71)           # gold AI spark
SPARK2 = (120, 200, 255)         # blue accent


def rounded(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def render(size: int) -> Image.Image:
    S = 512                       # render big, downscale later
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = int(S * 0.07)
    tile = (pad, pad, S - pad, S - pad)

    # Vertical-ish gradient tile (two stacked rounded rects for a subtle blend).
    rounded(d, tile, radius=int(S * 0.18), fill=EXCEL_GREEN2)
    rounded(d, (pad, pad, S - pad, int(S * 0.62)), int(S * 0.18), EXCEL_GREEN)

    # Spreadsheet grid (3 cols x 3 rows) inside the tile.
    gx0, gy0, gx1, gy1 = int(S * 0.20), int(S * 0.22), int(S * 0.80), int(S * 0.80)
    lw = max(3, int(S * 0.012))
    # outer frame
    d.rounded_rectangle((gx0, gy0, gx1, gy1), radius=int(S * 0.03),
                        outline=GRID, width=lw)
    # inner lines
    for k in (1, 2):
        x = gx0 + (gx1 - gx0) * k // 3
        d.line((x, gy0, x, gy1), fill=GRID, width=lw)
        y = gy0 + (gy1 - gy0) * k // 3
        d.line((gx0, y, gx1, y), fill=GRID, width=lw)
    # header row tint
    d.rectangle((gx0 + lw, gy0 + lw, gx1 - lw,
                 gy0 + (gy1 - gy0) // 3 - lw // 2),
                fill=(255, 255, 255, 60))

    # AI spark (four-point sparkle) top-right, suggesting 'intelligence'.
    cx, cy, r = int(S * 0.74), int(S * 0.26), int(S * 0.11)
    d.polygon([(cx, cy - r), (cx + r * 0.28, cy - r * 0.28),
               (cx + r, cy), (cx + r * 0.28, cy + r * 0.28),
               (cx, cy + r), (cx - r * 0.28, cy + r * 0.28),
               (cx - r, cy), (cx - r * 0.28, cy - r * 0.28)], fill=SPARK)
    # small secondary spark
    cx2, cy2, r2 = int(S * 0.86), int(S * 0.40), int(S * 0.045)
    d.polygon([(cx2, cy2 - r2), (cx2 + r2 * 0.3, cy2 - r2 * 0.3),
               (cx2 + r2, cy2), (cx2 + r2 * 0.3, cy2 + r2 * 0.3),
               (cx2, cy2 + r2), (cx2 - r2 * 0.3, cy2 + r2 * 0.3),
               (cx2 - r2, cy2), (cx2 - r2 * 0.3, cy2 - r2 * 0.3)], fill=SPARK2)

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    sizes = [16, 24, 32, 48, 64, 128, 256]
    base = render(256)
    base.save(OUT, format="ICO",
              sizes=[(s, s) for s in sizes])
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
