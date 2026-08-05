#!/usr/bin/env python3
"""Generate a dark-navy dashed arrow PNG with transparent background."""

from pathlib import Path
from PIL import Image, ImageDraw

NAVY = (11, 42, 74, 255)
W, H = 420, 160


def generate(path: Path) -> None:
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    y = H // 2
    x0, x1 = 28, 340
    thickness = 10
    dash_len, gap_len = 22, 14
    x = x0
    while x < x1:
        seg_end = min(x + dash_len, x1)
        r = thickness / 2
        draw.rounded_rectangle([x, y - r, seg_end, y + r], radius=r, fill=NAVY)
        x = seg_end + gap_len
    head_len, head_half_w = 48, 22
    tip = (x1 + head_len - 8, y)
    draw.polygon([tip, (x1 - 4, y - head_half_w), (x1 - 4, y + head_half_w)], fill=NAVY)
    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path, "PNG")
    print(f"Wrote {path}")


if __name__ == "__main__":
    out = Path(__file__).resolve().parent / "flecha_discontinua_azul_marino.png"
    generate(out)
    art = Path("/opt/cursor/artifacts/assets/flecha_discontinua_azul_marino.png")
    generate(art)
