"""Tạo logo tự thiết kế cho nhà thầu (Công ty TNHH Giang Long) dùng trong checklist."""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

S = 4  # siêu mẫu
W, H = 136 * S, 121 * S
im = Image.new("RGBA", (W, H), (255, 255, 255, 0))
d = ImageDraw.Draw(im)
bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F = lambda n: ImageFont.truetype(bold, n * S)
navy, gold = (18, 62, 112), (214, 150, 20)
cx, cy, r = W // 2, 46 * S, 36 * S
d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=navy)
d.ellipse([cx - r + 4 * S, cy - r + 4 * S, cx + r - 4 * S, cy + r - 4 * S], outline=gold, width=2 * S)
# hình kè chắn đất: tường bậc + sóng nước
d.polygon([(cx - 22 * S, cy + 16 * S), (cx - 22 * S, cy - 2 * S), (cx - 8 * S, cy - 2 * S), (cx - 8 * S, cy - 14 * S),
           (cx + 6 * S, cy - 14 * S), (cx + 6 * S, cy - 24 * S), (cx + 22 * S, cy - 24 * S), (cx + 22 * S, cy + 16 * S)], fill=(240, 240, 240))
for i, dx in enumerate(range(-22, 22, 11)):
    d.arc([cx + dx * S, cy + 8 * S, cx + (dx + 11) * S, cy + 20 * S], 180, 360, fill=gold, width=2 * S)
for text, y, n, col in [("GIANG LONG", 88, 15, navy), ("CÔNG TY TNHH", 108, 9, gold)]:
    w = d.textlength(text, font=F(n))
    d.text(((W - w) / 2, y * S - n * S // 2), text, font=F(n), fill=col)
out = Path(__file__).parent.parent / "data/qaqc/templates/logo_giang_long.png"
im.resize((136, 121), Image.LANCZOS).save(out)
print(out)
