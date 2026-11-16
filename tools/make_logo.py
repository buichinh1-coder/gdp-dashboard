"""Ô logo nhà thầu trong checklist: chỉ ghi tên công ty (không dùng hình logo)."""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

S = 4  # siêu mẫu
W, H = 136 * S, 121 * S
im = Image.new("RGBA", (W, H), (255, 255, 255, 0))
d = ImageDraw.Draw(im)
bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F = lambda n: ImageFont.truetype(bold, n * S)
navy = (18, 62, 112)
lines = [("CÔNG TY TNHH", 13), ("GIANG LONG", 19)]
gap = 8 * S
heights = [n * S for _, n in lines]
y = (H - (sum(heights) + gap * (len(lines) - 1))) // 2
for (text, n), h in zip(lines, heights):
    w = d.textlength(text, font=F(n))
    d.text(((W - w) / 2, y), text, font=F(n), fill=navy)
    y += h + gap
out = Path(__file__).parent.parent / "data/qaqc/templates/logo_giang_long.png"
im.resize((136, 121), Image.LANCZOS).save(out)
print(out)
