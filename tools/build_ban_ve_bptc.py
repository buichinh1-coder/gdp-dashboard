"""Bản vẽ biện pháp thi công (sơ đồ minh họa, A3 ngang, đen trắng) - 6 tờ.
Không có bản vẽ thiết kế nên không ghi kích thước; mọi số đo ghi 'theo hồ sơ thiết kế'.
Dùng: python tools/build_ban_ve_bptc.py <file_ra.pdf>
"""
import json
import math
import sys
from pathlib import Path

from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).parent.parent
PRJ = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))["project"]
FD = "/usr/share/fonts/truetype/liberation/LiberationSerif-"
for n, f in (("R", "Regular"), ("B", "Bold"), ("I", "Italic")):
    pdfmetrics.registerFont(TTFont("TNR-" + n, FD + f + ".ttf"))
mm = 72 / 25.4
W, H = landscape(A3)
NSHEET = 6


# ---------------------------------------------------------------- tiện ích vẽ
def M(v):
    return v * mm


def txt(c, x, y, s, size=9, font="R", anchor="l", rot=0):
    c.saveState()
    c.setFont("TNR-" + font, size)
    c.translate(M(x), M(y))
    if rot:
        c.rotate(rot)
    {"l": c.drawString, "c": c.drawCentredString, "r": c.drawRightString}[anchor](0, 0, s)
    c.restoreState()


def line(c, x1, y1, x2, y2, w=0.5, dash=None, gray=0):
    c.saveState()
    c.setLineWidth(w)
    c.setStrokeGray(gray)
    if dash:
        c.setDash(list(dash))
    c.line(M(x1), M(y1), M(x2), M(y2))
    c.restoreState()


def poly(c, pts, fill=None, w=0.6, close=True, dash=None, stroke=1):
    c.saveState()
    c.setLineWidth(w)
    if dash:
        c.setDash(list(dash))
    p = c.beginPath()
    p.moveTo(M(pts[0][0]), M(pts[0][1]))
    for x, y in pts[1:]:
        p.lineTo(M(x), M(y))
    if close:
        p.close()
    if fill is not None:
        c.setFillGray(fill)
    c.drawPath(p, fill=0 if fill is None else 1, stroke=stroke)
    c.restoreState()


def rect(c, x, y, w, h, fill=None, lw=0.6, dash=None):
    poly(c, [(x, y), (x + w, y), (x + w, y + h), (x, y + h)], fill=fill, w=lw, dash=dash)


def hatch(c, pts, kind="diag", step=2.2, gray=0.25):
    """Tô vật liệu: diag (đất/bê tông cốt thép), dots (bê tông), rock (đá xô), cross."""
    c.saveState()
    p = c.beginPath()
    p.moveTo(M(pts[0][0]), M(pts[0][1]))
    for x, y in pts[1:]:
        p.lineTo(M(x), M(y))
    p.close()
    c.clipPath(p, stroke=0, fill=0)
    xs, ys = [q[0] for q in pts], [q[1] for q in pts]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    c.setStrokeGray(gray)
    c.setFillGray(gray)
    c.setLineWidth(0.3)
    if kind in ("diag", "cross"):
        t = x0 - (y1 - y0)
        while t < x1:
            c.line(M(t), M(y0), M(t + (y1 - y0)), M(y1))
            t += step
        if kind == "cross":
            t = x0
            while t < x1 + (y1 - y0):
                c.line(M(t), M(y0), M(t - (y1 - y0)), M(y1))
                t += step
    elif kind == "dots":
        yy = y0
        k = 0
        while yy < y1:
            xx = x0 + (step / 2 if k % 2 else 0)
            while xx < x1:
                c.circle(M(xx), M(yy), 0.18 * mm, stroke=0, fill=1)
                xx += step
            yy += step * 0.8
            k += 1
    elif kind == "rock":
        yy = y0
        k = 0
        while yy < y1:
            xx = x0 + (step if k % 2 else 0)
            while xx < x1:
                c.circle(M(xx), M(yy), 0.9 * mm, stroke=1, fill=0)
                xx += step * 1.9
            yy += step * 1.4
            k += 1
    elif kind == "water":
        yy = y0 + 1.5
        while yy < y1:
            xx = x0
            while xx < x1:
                c.line(M(xx), M(yy), M(xx + 1.6), M(yy + 0.6))
                c.line(M(xx + 1.6), M(yy + 0.6), M(xx + 3.2), M(yy))
                xx += 3.2
            yy += 3.5
    c.restoreState()


def arrow(c, x1, y1, x2, y2, w=0.7, size=2.2, dash=None):
    line(c, x1, y1, x2 - 0.5 * math.cos(math.atan2(y2 - y1, x2 - x1)), y2 - 0.5 * math.sin(math.atan2(y2 - y1, x2 - x1)), w, dash)
    a = math.atan2(y2 - y1, x2 - x1)
    poly(c, [(x2, y2), (x2 - size * math.cos(a - 0.35), y2 - size * math.sin(a - 0.35)), (x2 - size * math.cos(a + 0.35), y2 - size * math.sin(a + 0.35))], fill=0, w=0.3)


def callout(c, x, y, n, size=4.0):
    c.saveState()
    c.setLineWidth(0.7)
    c.setFillGray(1)
    c.circle(M(x), M(y), M(size), stroke=1, fill=1)
    c.restoreState()
    txt(c, x, y - 1.1, str(n), 9, "B", "c")


def leader(c, x1, y1, x2, y2, text, anchor="l", size=9):
    line(c, x1, y1, x2, y2, 0.4)
    c.saveState()
    c.setFillGray(0)
    c.circle(M(x1), M(y1), 0.5 * mm, stroke=0, fill=1)
    c.restoreState()
    txt(c, x2 + (1.5 if anchor == "l" else -1.5), y2 - 1, text, size, "R", anchor)


def box(c, x, y, w, h, lines, size=9, bold_first=True, fill=None, lw=0.8):
    rect(c, x, y, w, h, fill=fill, lw=lw)
    n = len(lines)
    for i, s in enumerate(lines):
        txt(c, x + w / 2, y + h / 2 + (n - 1) * 2.4 - i * 4.8 - 1.2, s, size, "B" if (i == 0 and bold_first) else "R", "c")


def dim_h(c, x1, x2, y, text):
    arrow(c, (x1 + x2) / 2, y, x1, y, 0.4, 1.6)
    arrow(c, (x1 + x2) / 2, y, x2, y, 0.4, 1.6)
    line(c, x1, y - 2, x1, y + 2, 0.4)
    line(c, x2, y - 2, x2, y + 2, 0.4)
    c.saveState()
    c.setFillGray(1)
    c.rect(M((x1 + x2) / 2 - len(text) * 0.95), M(y - 1.8), M(len(text) * 1.9), M(3.6), stroke=0, fill=1)
    c.restoreState()
    txt(c, (x1 + x2) / 2, y - 1.1, text, 8.5, "I", "c")


def frame(c, no, title, sub=""):
    c.setLineWidth(1.4)
    c.rect(M(8), M(8), W - M(16), H - M(16))
    c.setLineWidth(0.4)
    c.rect(M(9.6), M(9.6), W - M(19.2), H - M(19.2))
    # khung tên (dưới)
    X0, X1, Y0, Y1 = 10, 410, 10, 50
    rect(c, X0, Y0, X1 - X0, Y1 - Y0, lw=1.0)
    cols = [X0, 105, 225, 275, 325, X1]
    for x in cols[1:-1]:
        line(c, x, Y0, x, Y1, 0.8)
    # ô 1: công trình
    txt(c, 57.5, 44, "CÔNG TRÌNH", 8.5, "B", "c")
    for i, s in enumerate(_wrap(PRJ["goi_thau"], 50)):
        txt(c, 57.5, 38.5 - i * 4.4, s, 9, "R", "c")
    txt(c, 57.5, 22, "Chủ đầu tư: " + PRJ["chu_dau_tu"], 8, "R", "c")
    txt(c, 57.5, 17.8, "Nhà thầu: " + PRJ["nha_thau"], 8, "R", "c")
    txt(c, 57.5, 13, "HĐ " + PRJ["so_hop_dong"], 7.5, "I", "c")
    # ô 2: tên bản vẽ
    txt(c, 165, 44, "BẢN VẼ BIỆN PHÁP THI CÔNG", 9, "B", "c")
    for i, s in enumerate(_wrap(title.upper(), 36)):
        txt(c, 165, 36 - i * 6, s, 13, "B", "c")
    if sub:
        txt(c, 165, 15, sub, 8.5, "I", "c")
    # ô 3: số hiệu
    txt(c, 250, 44, "SỐ HIỆU", 8.5, "B", "c")
    txt(c, 250, 34, f"BV-{no:02d}", 18, "B", "c")
    txt(c, 250, 24, f"Tờ {no}/{NSHEET}", 10, "R", "c")
    txt(c, 250, 15, "Tỉ lệ: KHÔNG TỈ LỆ", 9, "R", "c")
    # ô 4: lập / ký
    for i, (k, v) in enumerate([("Người lập", ""), ("Kiểm tra", ""), ("Phê duyệt", "")]):
        yy = 44 - i * 12
        txt(c, 277, yy, k + ":", 8.5, "B")
        if v:
            txt(c, 298, yy, v, 8.5, "R")
        line(c, 276, yy - 3, 323, yy - 3, 0.3)
    # ô 5: ký duyệt các bên
    for i, k in enumerate(["CHỦ ĐẦU TƯ", "TVQLDA & GIÁM SÁT", "NHÀ THẦU"]):
        xx = 325 + (X1 - 325) * (i + 0.5) / 3
        txt(c, xx, 44, k, 6.6, "B", "c")
        txt(c, xx, 12.5, "(Ký, đóng dấu)", 6.5, "I", "c")
        line(c, 325 + (X1 - 325) * i / 3, Y0, 325 + (X1 - 325) * i / 3, Y1, 0.4) if i else None
    # ghi chú chung
    txt(c, 15, 54, "Ghi chú: Bản vẽ là sơ đồ minh họa biện pháp thi công, KHÔNG THEO TỈ LỆ. Kích thước, cao độ, quy cách vật liệu lấy theo hồ sơ thiết kế và chỉ dẫn kỹ thuật được duyệt.", 8.5, "I")


def _wrap(s, n):
    out, cur = [], ""
    for w_ in s.split():
        if len(cur) + len(w_) + 1 > n:
            out.append(cur)
            cur = w_
        else:
            cur = (cur + " " + w_).strip()
    return out + [cur]


def title_area(c, s):
    txt(c, 210, 281, s, 12, "B", "c")


# ---------------------------------------------------------------- BV-01 mặt bằng
def bv01(c):
    frame(c, 1, "Sơ đồ mặt bằng tổ chức thi công", "Tuyến 3, đoạn cọc 88 – cọc 126")
    # vùng vẽ y 60..280
    title_area(c, "SƠ ĐỒ MẶT BẰNG TỔ CHỨC THI CÔNG TOÀN TUYẾN (cọc 88 – cọc 126)")
    x0, x1 = 35, 385
    n = 4
    wseg = (x1 - x0) / n
    yk = 185  # trục kè
    # lạch phía núi
    poly(c, [(x0 - 10, yk + 28), (x1 + 10, yk + 28), (x1 + 10, yk + 52), (x0 - 10, yk + 52)])
    hatch(c, [(x0 - 10, yk + 28), (x1 + 10, yk + 28), (x1 + 10, yk + 52), (x0 - 10, yk + 52)], "water")
    txt(c, 210, yk + 56, "LẠCH PHÍA NÚI (đào bùn, nạo vét mở mặt bằng; hoàn thiện sau)", 10, "B", "c")
    # núi
    txt(c, 210, yk + 78, "PHÍA NÚI", 11, "B", "c")
    for i in range(0, 14):
        xx = 40 + i * 26
        poly(c, [(xx, yk + 62), (xx + 10, yk + 70), (xx + 20, yk + 62)], w=0.4, close=False)
    # sàn thi công + đê quai (dải)
    poly(c, [(x0 - 10, yk + 8), (x1 + 10, yk + 8), (x1 + 10, yk + 28), (x0 - 10, yk + 28)], fill=0.92, w=0.5)
    txt(c, 210, yk + 19, "ĐÊ QUAI, SÀN THI CÔNG VÀ ĐƯỜNG CÔNG VỤ (đắp K95, giữ thông xe máy thi công)", 9, "R", "c")
    # tuyến thoát nước tạm
    line(c, x0 - 10, yk + 11, x1 + 10, yk + 11, 0.9, (4, 2))
    txt(c, x0 - 7, yk + 12.6, "Rãnh thoát nước tạm (nét đứt)", 8, "I")
    # kè
    poly(c, [(x0, yk - 5), (x1, yk - 5), (x1, yk + 5), (x0, yk + 5)], fill=0.55, w=1.0)
    txt(c, 210, yk - 1.2, "TUYẾN KÈ BÊ TÔNG CỐT THÉP ĐỔ TẠI CHỖ", 9, "B", "c")
    # phía ngoài
    poly(c, [(x0 - 10, yk - 45), (x1 + 10, yk - 45), (x1 + 10, yk - 5), (x0 - 10, yk - 5)], w=0.4, dash=(2, 2))
    txt(c, 210, yk - 40, "PHÍA NGOÀI KÈ (đá xô chân tường chắn)", 9, "R", "c")
    for i in range(n):
        xa = x0 + i * wseg
        if i:
            line(c, xa, yk - 48, xa, yk + 52, 0.9, (6, 2))
        # nhãn phân đoạn
        box(c, xa + 8, yk - 32, wseg - 16, 12, [f"PHÂN ĐOẠN {i + 1}"], 10, True, fill=1)
        # cọc đầu
    cocs = [88, 98, 108, 118, 126]
    for i, cc in enumerate(cocs):
        xa = x0 + i * wseg if i < 4 else x1
        line(c, xa, yk - 10, xa, yk - 14, 0.9)
        txt(c, xa, yk - 19, f"Cọc {cc}", 9, "B", "c")
    # hướng thi công
    arrow(c, x0 + 15, yk - 56, x1 - 15, yk - 56, 1.2, 3)
    txt(c, 210, yk - 53, "HƯỚNG THI CÔNG CUỐN CHIẾU: PĐ1 → PĐ2 → PĐ3 → PĐ4 (đê quai → nạo vét → cọc → móng → kết cấu → đắp hoàn trả)", 9, "R", "c")
    # bố trí phụ trợ (phía dưới)
    ya = 88
    box(c, 20, ya, 62, 30, ["VĂN PHÒNG CHỈ HUY", "Nhà vệ sinh, y tế"], 9)
    box(c, 90, ya, 62, 30, ["KHO VẬT TƯ", "Xi măng phụ gia, dụng cụ"], 9)
    box(c, 160, ya, 62, 30, ["BÃI GIA CÔNG CỐT THÉP", "VÀ VÁN KHUÔN"], 9)
    box(c, 230, ya, 62, 30, ["BÃI TẬP KẾT CỌC BTCT", "300x300 (đạt R7)"], 9)
    box(c, 300, ya, 55, 30, ["BÃI TẠM ĐẤT", "ĐÁ XÔ"], 9)
    box(c, 362, ya, 40, 30, ["CỔNG RA VÀO", "RỬA XE"], 9)
    txt(c, 210, ya + 34, "CÁC KHU PHỤ TRỢ (vị trí cụ thể theo hiện trạng mặt bằng được CĐT bàn giao)", 10, "B", "c")
    for xx in (51, 121, 191, 261, 327, 382):
        arrow(c, xx, ya + 30, xx, yk - 52, 0.5, 2, (2, 2)) if False else None
    # chú dẫn
    txt(c, 20, 78, "Chú dẫn: ", 9, "B")
    txt(c, 40, 78, "- - -  ranh giới phân đoạn      ▬▬  tuyến kè      ┅  rãnh thoát nước tạm      ≈  lạch/nước      Ngưỡng nối ra đường công cộng qua cổng rửa xe.", 9, "R")
    txt(c, 20, 71, "Hạ tầng tạm: bơm thoát nước hố móng, biển báo, rào chắn, chiếu sáng ban đêm (nếu làm đêm), thùng rác và hố lắng tại khu rửa xe (xem thuyết minh).", 9, "I")


# ---------------------------------------------------------------- BV-02 mặt cắt
def bv02(c):
    frame(c, 2, "Mặt cắt ngang điển hình thi công kè", "Mặt cắt A-A (sơ đồ)")
    title_area(c, "MẶT CẮT NGANG ĐIỂN HÌNH TRONG QUÁ TRÌNH THI CÔNG VÀ HOÀN THIỆN")
    G = 175  # mặt đất tự nhiên
    xl, xr = 25, 395
    # nền đất
    poly(c, [(xl, G), (xr, G), (xr, 80), (xl, 80)], w=0.6)
    hatch(c, [(xl, G), (xr, G), (xr, 80), (xl, 80)], "diag", 3.2, 0.5)
    # hố móng đào
    pit = [(130, G), (150, G - 42), (250, G - 42), (270, G)]
    poly(c, [(125, G + 0.3), (275, G + 0.3), (270, G), (250, G - 42), (150, G - 42), (130, G)], fill=1, w=0.8)
    # đê quai & sàn thi công hai bên hố (đắp K95)
    poly(c, [(60, G), (60, G + 12), (122, G + 12), (135, G), ], fill=0.9, w=0.7)  # trái
    poly(c, [(265, G), (278, G + 12), (350, G + 12), (350, G)], fill=0.9, w=0.7)  # phải
    txt(c, 90, G + 15, "Đê quai, sàn thi công (K95)", 9, "R", "c")
    txt(c, 315, G + 15, "Đê quai, sàn thi công (K95)", 9, "R", "c")
    # lạch phía núi (phải)
    poly(c, [(350, G), (395, G), (395, G - 14), (365, G - 18), (350, G - 5)], fill=1, w=0.7)
    hatch(c, [(352, G - 3), (394, G - 3), (394, G - 13), (365, G - 17), (352, G - 6)], "water")
    txt(c, 372, G - 24, "Lạch phía núi", 9, "I", "c")
    txt(c, 372, G - 29, "(nạo vét mở mặt bằng)", 8.5, "I", "c")
    # cọc
    for i in range(6):
        px = 158 + i * 17
        rect(c, px - 2.2, G - 42 - 34, 4.4, 34 + 4, fill=0.7, lw=0.7)
    # lót
    rect(c, 140, G - 42, 120, 3, fill=0.88, lw=0.6)
    hatch(c, [(140, G - 42), (260, G - 42), (260, G - 39), (140, G - 39)], "dots", 1.6)
    # móng
    mp = [(140, G - 39), (260, G - 39), (260, G - 27), (140, G - 27)]
    poly(c, mp, fill=1, w=0.9)
    hatch(c, mp, "dots", 2.0, 0.3)
    # thân (bản đứng) - lệch trái
    th = [(168, G - 27), (186, G - 27), (186, G + 10), (168, G + 10)]
    poly(c, th, fill=1, w=0.9)
    hatch(c, th, "dots", 2.0, 0.3)
    kd = [(163, G + 10), (191, G + 10), (191, G + 18), (163, G + 18)]
    poly(c, kd, fill=1, w=0.9)
    hatch(c, kd, "dots", 2.0, 0.3)
    # đất đắp phía trong (phải của thân) , đá xô phía ngoài (trái)
    inn = [(186, G - 27), (260, G - 27), (270, G), (275, G + 12) if False else (270, G), (186, G + 10)]
    inn = [(186, G - 27), (260, G - 27), (270, G), (186, G + 10)]
    poly(c, inn, fill=1, w=0.6)
    hatch(c, inn, "diag", 3.0, 0.45)
    for k in range(1, 5):  # lớp đắp
        yy = -27 + k * 8
        line(c, 186, G + yy, 258 + k * 2.5, G + yy, 0.3, (2, 1.5), 0.4)
    out = [(130, G), (150, G - 42), (140, G - 39) if False else (150, G - 42), (168, G - 27), (168, G + 10), (150, G + 10)] if False else [(140, G - 27), (168, G - 27), (168, G + 10), (148, G + 10), (132, G)]
    poly(c, [(140, G - 27), (168, G - 27), (168, G + 6), (150, G + 6), (132, G - 4)], fill=1, w=0.6)
    hatch(c, [(140, G - 27), (168, G - 27), (168, G + 6), (150, G + 6), (132, G - 4)], "rock", 3.0, 0.35)
    # nhãn có số
    L = [(1, "Đê quai, sàn thi công, rãnh thoát nước tạm (K95)", (90, G + 6), (35, 215)),
         (2, "Đào bùn, nạo vét lạch phía núi", (370, G - 10), (320, 225)),
         (3, "Đào móng bằng máy (đất cấp II) theo từng đốt", (260, G - 10), (270, 215)),
         (4, "Cọc BTCT 300x300 M300 bền sunfat – ép đến độ chối/lực ép thiết kế", (176, G - 60), (40, 110)),
         (5, "Cắt, đập đầu cọc, neo cốt thép chờ vào móng", (150, G - 36), (30, 128)),
         (6, "Bê tông lót M100 (ván khuôn lót)", (225, G - 41), (275, 100)),
         (7, "Móng, thân, khóa đỉnh BTCT M300 bền sunfat B12", (230, G - 33), (285, 118)),
         (8, "Đắp hoàn trả hố móng phía trong kè, từng lớp, đầm K95", (235, G - 12), (292, 135)),
         (9, "Đá xô chân tường chắn phía ngoài kè", (152, G - 12), (30, 146))]
    for n, t, (px, py), (tx, ty) in L:
        callout(c, tx if n in (1, 4, 5, 9) else tx, ty, n) if False else None
    # chú thích bằng danh sách số bên phải/ trái
    ylist = 262
    txt(c, 25, 275, "CÁC BƯỚC THI CÔNG VÀ HẠNG MỤC TRÊN MẶT CẮT", 10, "B")
    for n, t, (px, py), _ in L:
        yy = ylist - (n - 1) * 5.8
        callout(c, 29, yy + 1, n, 2.6)
        txt(c, 35, yy, t, 9)
        callout(c, px, py, n, 2.6)
    # ghi chú hướng
    txt(c, 100, G - 80, "PHÍA NGOÀI KÈ", 10, "B", "c")
    txt(c, 330, G - 80, "PHÍA TRONG KÈ (phía đất, hướng núi)", 10, "B", "c")
    dim_h(c, 140, 260, G - 90, "Bề rộng móng, chiều cao tường, cao độ: theo thiết kế")
    txt(c, 210, 62 + 0, "Hướng \"phía trong/phía ngoài\" mô tả theo tên hạng mục trong lịch thi công; bố trí cụ thể theo bản vẽ thiết kế.", 8.5, "I", "c")


# ---------------------------------------------------------------- BV-03 sơ đồ trình tự
def bv03(c):
    frame(c, 3, "Sơ đồ trình tự thi công và các điểm nghiệm thu", "Cho mỗi phân đoạn")
    title_area(c, "SƠ ĐỒ TRÌNH TỰ THI CÔNG MỘT PHÂN ĐOẠN VÀ CÁC ĐIỂM DỪNG NGHIỆM THU")
    steps = [("1", "Đắp đê quai, tạo sàn thi công, thoát nước tạm", "NT công việc (K95)"),
             ("2", "Đào bùn, nạo vét lạch phía núi (đợt mở mặt bằng)", "NT công việc"),
             ("3", "Cọc BTCT 300x300: nghiệm thu vật liệu đầu vào (R7)", "NT vật liệu đầu vào"),
             ("4", "Đào móng bằng máy, vận chuyển đất đào", "NT công việc"),
             ("5", "Ép cọc, cắt đầu cọc", "NT công việc"),
             ("6", "Ván khuôn và đổ bê tông lót M100", "NT công việc"),
             ("7", "Cốt thép, ván khuôn móng/thân/khóa đỉnh, khe lún", "NT trước khi đổ BT"),
             ("8", "Đổ bê tông M300 bền sunfat, lấy mẫu R7/R28", "NT công việc + lấy mẫu"),
             ("9", "Đắp hoàn trả hố móng phía trong kè, đầm K95", "NT từng lớp"),
             ("10", "Đắp đá xô chân tường chắn phía ngoài kè", "NT công việc"),
             ("11", "Hoàn thiện nạo vét lạch phía núi (30% còn lại)", "NT công việc")]
    xL, wL = 20, 200
    y = 270
    hh = 14.5
    for i, (n, t, g) in enumerate(steps):
        yy = y - i * 17
        box(c, xL, yy - hh, wL, hh, [], 9)
        callout(c, xL + 8, yy - hh / 2, n, 4.2)
        txt(c, xL + 16, yy - hh / 2 - 1.3, t, 10.5, "B")
        # điểm nghiệm thu
        poly(c, [(xL + wL + 6, yy - hh / 2), (xL + wL + 16, yy - hh / 2 + 6), (xL + wL + 26, yy - hh / 2), (xL + wL + 16, yy - hh / 2 - 6)], fill=0.85, w=0.8)
        txt(c, xL + wL + 16, yy - hh / 2 - 1, "NT", 8.5, "B", "c")
        line(c, xL + wL, yy - hh / 2, xL + wL + 6, yy - hh / 2, 0.7)
        txt(c, xL + wL + 30, yy - hh / 2 - 1.2, g, 9, "I")
        if i < len(steps) - 1:
            arrow(c, xL + 8, yy - hh, xL + 8, yy - 17 + 0.1, 0.8, 2.2)
    # bên phải: cuốn chiếu
    xr = 305
    txt(c, xr + 45, 270, "TỔ CHỨC CUỐN CHIẾU THEO PHÂN ĐOẠN", 10.5, "B", "c")
    rows = [("PĐ1", "Cọc 88–98", "14/08 – 12/09"), ("PĐ2", "Cọc 98–108", "19/08 – 17/09"), ("PĐ3", "Cọc 108–118", "24/08 – 22/09"), ("PĐ4", "Cọc 118–126", "29/08 – 27/09")]
    ncols = [("Phân đoạn", 18), ("Đoạn", 24), ("Hoàn thành kết cấu BT*", 44)]
    yy = 262
    xx = xr
    for t, w in ncols:
        rect(c, xx, yy - 8, w, 8, fill=0.85)
        txt(c, xx + w / 2, yy - 5.6, t, 8, "B", "c")
        xx += w
    for i, r in enumerate(rows):
        xx = xr
        for (t, w), v in zip(ncols, r):
            rect(c, xx, yy - 16 - i * 8, w, 8)
            txt(c, xx + w / 2, yy - 13.6 - i * 8, v, 8.5, "R", "c")
            xx += w
    txt(c, xr, yy - 52, "* Khoảng từ ngày bắt đầu việc đầu tiên đến ngày hoàn thành đổ", 8, "I")
    txt(c, xr, yy - 56, "   bê tông M300 theo lịch nghiệm thu đã xác nhận.", 8, "I")
    # chú dẫn
    txt(c, xr, 195, "NGUYÊN TẮC", 10.5, "B")
    for i, s in enumerate(["- Mỗi phân đoạn chỉ chuyển sang bước sau khi bước trước", "  đã nghiệm thu (checklist + biên bản).",
                           "- Phân đoạn sau bắt đầu khi phân đoạn trước đã có", "  mặt bằng và máy móc dịch chuyển (cuốn chiếu).",
                           "- Cọc: chỉ ép khi đạt R7 và hồ sơ đầu vào được chấp thuận.",
                           "- Bê tông: lấy mẫu R7, R28 theo từng đợt đổ.",
                           "- Đắp hoàn trả: nghiệm thu độ chặt K95 từng lớp,",
                           "  kiểm tra tại hiện trường theo từng ngày đắp.",
                           "- Hồ sơ nghiệm thu lập theo mẫu Apave trong hồ sơ", "  chất lượng kèm theo."]):
        txt(c, xr, 187 - i * 5.4, s, 9)
    box(c, xr, 110, 100, 20, ["ĐƯỜNG GĂNG: chuỗi công việc phân đoạn 4", "(xem Tiến độ thi công)"], 9, True, fill=0.92)


# ---------------------------------------------------------------- BV-04 ép cọc
def bv04(c):
    frame(c, 4, "Biện pháp thi công ép cọc BTCT 300x300", "Mặt đứng và mặt bằng ép cọc (sơ đồ)")
    title_area(c, "BIỆN PHÁP THI CÔNG ÉP CỌC VÀ CẮT ĐẦU CỌC")
    G = 175
    # nền
    poly(c, [(25, G), (265, G), (265, 70), (25, 70)])
    hatch(c, [(25, G), (265, G), (265, 70), (25, 70)], "diag", 3.2, 0.55)
    # cọc đã ép và cọc đang ép
    for i, px in enumerate((70, 90, 110)):
        rect(c, px - 3, G - 62, 6, 62 + 0, fill=0.75, lw=0.7)
        rect(c, px - 3, G, 6, 4, fill=0.75, lw=0.7)
    txt(c, 80, G - 69, "Cọc đã ép", 8.5, "I", "c")
    # cọc đang ép
    rect(c, 165 - 3, G - 20, 6, 20 + 24, fill=0.75, lw=0.7)
    rect(c, 165 - 3, G + 24, 6, 20, fill=1, lw=0.7)
    # khung ép
    rect(c, 130, G + 44, 70, 5, fill=0.45, lw=0.8)
    rect(c, 128, G, 4, 49, fill=0.55, lw=0.7)
    rect(c, 198, G, 4, 49, fill=0.55, lw=0.7)
    rect(c, 145, G + 49, 40, 10, fill=0.7, lw=0.8)  # đầu kích
    rect(c, 160, G + 59, 10, 22, fill=0.85, lw=0.8)  # xi lanh
    # đối trọng
    rect(c, 125, G + 5, 80, 6, fill=1)
    rect(c, 118, G + 11, 94, 8, fill=0.3)
    rect(c, 122, G + 19, 86, 8, fill=0.3)
    txt(c, 165, G + 31, "Đối trọng / dầm dẫn (theo máy ép sử dụng)", 8.5, "R", "c")
    arrow(c, 165, G + 100, 165, G + 83, 1.2, 3)
    txt(c, 172, G + 92, "Lực ép Pép (theo thiết kế)", 9, "R")
    # cẩu/ máy
    rect(c, 215, G, 35, 22, fill=0.8, lw=0.8)
    poly(c, [(218, G + 22), (240, G + 22), (248, G + 70), (244, G + 70), (232, G + 28)], fill=0.9, w=0.6)
    line(c, 247, G + 70, 195, G + 40, 0.5)
    txt(c, 232, G + 76, "Cẩu đưa cọc vào vị trí", 8.5, "R", "c")
    # nhãn
    L = ["Cọc BTCT 300x300 M300 bền sunfat, đạt R7", "Khung dẫn hướng, định vị tim cọc, kiểm tra độ thẳng đứng", "Đồng hồ áp lực, ghi nhật ký ép từng cọc",
         "Dừng ép khi đạt lực ép và độ chối theo thiết kế", "Nối cọc (nếu thiết kế quy định): hàn đảm bảo thẳng trục"]
    for i, s in enumerate(L):
        callout(c, 30, 262 - i * 7 + 1, i + 1, 2.6)
        txt(c, 36, 262 - i * 7, s, 9)
    callout(c, 165, G - 36 if False else G - 10, 1, 2.6)
    callout(c, 200, G + 47, 2, 2.6)
    callout(c, 152, G + 66, 3, 2.6)
    callout(c, 165, G - 40, 4, 2.6)
    # mặt bằng thứ tự ép (phải)
    px0, py0 = 285, 170
    txt(c, 340, 270, "MẶT BẰNG BỐ TRÍ VÀ THỨ TỰ ÉP CỌC (sơ đồ)", 10, "B", "c")
    rect(c, px0, py0 - 10, 110, 80, lw=0.8, dash=(4, 2))
    n = 0
    for row in range(5):
        for col in range(6):
            n += 1
            cx, cy = px0 + 12 + col * 17.5, py0 + 62 - row * 14
            c.saveState()
            c.setLineWidth(0.6)
            c.rect(M(cx - 3), M(cy - 3), M(6), M(6))
            c.restoreState()
    arrow(c, px0 + 8, py0 + 78, px0 + 100, py0 + 78, 1.0, 2.5)
    txt(c, px0 + 55, py0 + 81, "Hướng ép: theo một phía, từ đoạn đã có mặt bằng", 8.5, "I", "c")
    txt(c, px0, py0 - 15, "Số lượng, khoảng cách, cách bố trí cọc trong móng: theo thiết kế.", 8.5, "I")
    # cắt đầu cọc detail
    txt(c, 340, 142, "CẮT, ĐẬP ĐẦU CỌC", 10, "B", "c")
    dx, dy = 300, 62
    rect(c, dx, dy, 14, 40, fill=0.75, lw=0.8)
    rect(c, dx, dy + 40, 14, 12, fill=1, lw=0.8, dash=(3, 2))
    for k in range(3):
        line(c, dx + 3 + k * 4, dy + 38, dx + 3 + k * 4, dy + 60, 0.7)
    line(c, dx - 12, dy + 40, dx + 40, dy + 40, 0.8, (6, 2, 1, 2))
    txt(c, dx + 22, dy + 39, "Cao độ cắt theo thiết kế", 9)
    txt(c, dx + 22, dy + 52, "Phần đầu cọc phá bỏ, xúc vận chuyển", 9)
    txt(c, dx + 22, dy + 62, "Cốt thép chờ neo vào móng (theo TK)", 9)
    txt(c, dx + 22, dy + 20, "Mặt cắt sau khi đập: phẳng, sạch", 9)


# ---------------------------------------------------------------- BV-05 móng, ván khuôn, đổ bê tông
def bv05(c):
    frame(c, 5, "Biện pháp thi công móng, ván khuôn và đổ bê tông", "Sơ đồ đổ bê tông và chi tiết khe lún")
    title_area(c, "BIỆN PHÁP ĐÀO MÓNG, VÁN KHUÔN, CỐT THÉP VÀ ĐỔ BÊ TÔNG M300")
    # trái: mặt cắt ván khuôn
    G = 200
    txt(c, 110, 272, "A. MẶT CẮT VÁN KHUÔN VÀ ĐỔ BÊ TÔNG", 10, "B", "c")
    poly(c, [(25, G), (200, G), (200, 80), (25, 80)])
    hatch(c, [(25, G), (200, G), (200, 80), (25, 80)], "diag", 3.2, 0.55)
    poly(c, [(40, G), (185, G), (185, G - 60), (40, G - 60)], fill=1)
    rect(c, 50, G - 60, 125, 3, fill=0.88)
    mp = [(55, G - 57), (170, G - 57), (170, G - 45), (55, G - 45)]
    poly(c, mp, fill=1, w=0.9)
    hatch(c, mp, "dots", 2.0, 0.3)
    for k in range(5):
        rect(c, 66 + k * 22, G - 70, 4.4, 14, fill=0.75, lw=0.6)
    th = [(95, G - 45), (113, G - 45), (113, G - 8), (95, G - 8)]
    poly(c, th, fill=1, w=0.9)
    hatch(c, th, "dots", 2.0, 0.3)
    # ván khuôn thân
    for xx in (92, 116):
        rect(c, xx - 1.8 if xx == 92 else xx - 0.4, G - 45, 2.2, 37 + 8, fill=0.4, lw=0.5)
    # chống xiên
    line(c, 92, G - 20, 72, G - 50, 0.8)
    line(c, 116, G - 20, 136, G - 50, 0.8)
    rect(c, 69, G - 52, 8, 3, fill=0.3)
    rect(c, 132, G - 52, 8, 3, fill=0.3)
    kd = [(90, G - 8), (118, G - 8), (118, G), (90, G)]
    poly(c, kd, fill=1, w=0.9)
    hatch(c, kd, "dots", 2.0, 0.3)
    # thép
    for xx in (98, 104, 110):
        line(c, xx, G - 44, xx, G - 2, 0.9, gray=0.2)
    line(c, 58, G - 52, 168, G - 52, 0.9, gray=0.2)
    line(c, 58, G - 48, 168, G - 48, 0.9, gray=0.2)
    # xe bơm / cần
    rect(c, 150, G + 8, 45, 16, fill=0.8, lw=0.8)
    line(c, 158, G + 24, 130, G + 40, 1.2)
    line(c, 130, G + 40, 106, G + 6, 1.2)
    arrow(c, 106, G + 12, 106, G - 2, 0.8, 2.2)
    txt(c, 172, G + 28, "Xe bơm bê tông / cần phân phối", 8.5, "R", "c")
    rect(c, 170, G + 28 + 6, 25, 10, fill=0.9, lw=0.8)
    txt(c, 182.5, G + 36, "Xe bồn BTTP", 7.5, "R", "c")
    for i, s in enumerate(["Bê tông lót M100 trên nền đã nghiệm thu", "Cốt thép móng, thân, khóa đỉnh: nghiệm thu trước khi đổ",
                           "Ván khuôn, cây chống, thanh giằng ổn định", "Đổ bê tông phân lớp, đầm dùi; độ dày lớp theo chỉ dẫn kỹ thuật",
                           "Lấy mẫu tại hiện trường theo từng đợt đổ (R7, R28)", "Bảo dưỡng ẩm theo TCVN 8828:2011"]):
        callout(c, 30, 74 - i * 0 + 0, 0, 0) if False else None
    for i, s in enumerate(["1. Bê tông lót M100 trên nền đã nghiệm thu", "2. Cốt thép móng, thân, khóa đỉnh: nghiệm thu trước khi đổ",
                           "3. Ván khuôn, cây chống, thanh giằng: ổn định, kín khít", "4. Đổ bê tông phân lớp, đầm dùi; độ dày lớp theo chỉ dẫn kỹ thuật",
                           "5. Lấy mẫu hiện trường theo từng đợt đổ (R7, R28)", "6. Bảo dưỡng ẩm theo TCVN 8828:2011"]):
        pass
    # trái dưới ghi chú
    # phải: khe lún
    txt(c, 305, 272, "B. CHI TIẾT KHE LÚN VÀ THOÁT NƯỚC (sơ đồ)", 10, "B", "c")
    kx, ky = 260, 190
    rect(c, kx, ky, 50, 60, fill=1, lw=0.9)
    hatch(c, [(kx, ky), (kx + 50, ky), (kx + 50, ky + 60), (kx, ky + 60)], "dots", 2.2, 0.3)
    rect(c, kx + 50, ky, 8, 60, fill=0.8, lw=0.9)
    rect(c, kx + 58, ky, 50, 60, fill=1, lw=0.9)
    hatch(c, [(kx + 58, ky), (kx + 108, ky), (kx + 108, ky + 60), (kx + 58, ky + 60)], "dots", 2.2, 0.3)
    txt(c, kx + 54, ky + 63, "Khe lún", 9, "B", "c")
    txt(c, kx + 54, ky - 5, "Vật liệu chèn khe, băng cản nước (theo thiết kế)", 8.5, "I", "c")
    txt(c, kx + 25, ky + 30, "Đốt N", 9, "B", "c")
    txt(c, kx + 83, ky + 30, "Đốt N+1", 9, "B", "c")
    # ống thoát nước
    c.saveState()
    c.setLineWidth(0.9)
    c.circle(M(kx + 25), M(ky + 12), M(3.2), stroke=1, fill=0)
    c.restoreState()
    leader(c, kx + 25, ky + 12, kx + 4, ky - 20, "Chi tiết thoát nước (ống, lỗ thoát) theo bản vẽ")
    # lấy mẫu
    txt(c, 305, 160, "C. KIỂM SOÁT CHẤT LƯỢNG KHI ĐỔ BÊ TÔNG", 10, "B", "c")
    for i, s in enumerate(["- Kiểm tra phiếu giao hàng, thời gian trộn – đổ, độ sụt theo cấp phối đã duyệt.",
                           "- Lấy mẫu tại hiện trường: mỗi tổ 3 viên, bảo dưỡng chuẩn, nén R7 và R28.",
                           "- Số tổ mẫu mỗi đợt đổ theo thể tích (xem biên bản lấy mẫu LM).",
                           "- Không bổ sung nước vào xe bồn tại hiện trường.",
                           "- Đầm dùi đủ thời gian, không chạm ván khuôn và cốt thép.",
                           "- Bảo dưỡng ẩm tối thiểu theo TCVN 8828:2011, tháo ván khuôn khi đạt cường độ."]):
        txt(c, 255, 150 - i * 6.5, s, 9)
    txt(c, 25, 72, "Trình tự mỗi đốt: nghiệm thu nền/lót → cốt thép → ván khuôn → đổ bê tông M300 → bảo dưỡng → tháo ván khuôn → nghiệm thu.", 9, "B")


# ---------------------------------------------------------------- BV-06 đắp hoàn trả, đá xô, an toàn
def bv06(c):
    frame(c, 6, "Biện pháp đắp hoàn trả, đá xô, an toàn và môi trường", "Sơ đồ đắp theo lớp và bố trí an toàn")
    title_area(c, "BIỆN PHÁP ĐẮP HOÀN TRẢ, ĐẮP ĐÁ XÔ, AN TOÀN LAO ĐỘNG VÀ VỆ SINH MÔI TRƯỜNG")
    txt(c, 110, 272, "A. ĐẮP HOÀN TRẢ HỐ MÓNG THEO LỚP (phía trong kè)", 10, "B", "c")
    G = 228
    poly(c, [(25, G - 70), (195, G - 70), (195, G - 100), (25, G - 100)])
    hatch(c, [(25, G - 70), (195, G - 70), (195, G - 100), (25, G - 100)], "diag", 3.2, 0.55)
    # tường
    w = [(60, G - 70), (78, G - 70), (78, G - 5), (60, G - 5)]
    poly(c, w, fill=1, w=0.9)
    hatch(c, w, "dots", 2.0, 0.3)
    for k in range(5):
        yy = G - 70 + k * 12 + 12
        poly(c, [(78, yy - 12), (190 - k * 5, yy - 12), (190 - k * 5, yy), (78, yy)], fill=0.98 - 0.03 * k, w=0.5)
        hatch(c, [(78, yy - 12), (190 - k * 5, yy - 12), (190 - k * 5, yy), (78, yy)], "diag", 3.0, 0.6)
        txt(c, 134 - k * 2.5, yy - 7, f"Lớp {k + 1}", 8.5, "R", "c")
    rect(c, 150, G - 12, 38, 12, fill=0.8, lw=0.8)
    rect(c, 154, G - 12 - 6, 30, 6, fill=0.5, lw=0.8)
    txt(c, 169, G + 4, "Lu / đầm", 8.5, "R", "c")
    # đá xô
    txt(c, 305, 272, "B. ĐẮP ĐÁ XÔ CHÂN TƯỜNG CHẮN (phía ngoài kè)", 10, "B", "c")
    ex, ey = 230, 228
    poly(c, [(ex, ey - 100), (ex + 150, ey - 100), (ex + 150, ey - 70), (ex, ey - 70)])
    hatch(c, [(ex, ey - 100), (ex + 150, ey - 100), (ex + 150, ey - 70), (ex, ey - 70)], "diag", 3.2, 0.55)
    wl = [(ex + 80, ey - 70), (ex + 98, ey - 70), (ex + 98, ey - 5), (ex + 80, ey - 5)]
    poly(c, wl, fill=1, w=0.9)
    hatch(c, wl, "dots", 2.0, 0.3)
    st = [(ex + 20, ey - 70), (ex + 80, ey - 70), (ex + 80, ey - 30), (ex + 60, ey - 28), (ex + 40, ey - 42)]
    poly(c, st, fill=1, w=0.8)
    hatch(c, st, "rock", 3.0, 0.35)
    txt(c, ex + 50, ey - 22, "Đá xô chân tường chắn", 9, "R", "c")
    txt(c, ex + 50, ey - 27, "(xếp/đổ theo thiết kế, tới cao độ quy định)", 8.5, "I", "c")
    # yêu cầu
    for i, s in enumerate(["- Đắp từng lớp, chiều dày lớp và số lượt lu theo chỉ dẫn kỹ thuật, đầm thử để xác định.", "- Kiểm tra độ chặt K95 tại hiện trường theo từng ngày đắp (phễu rót cát hoặc dao đai).",
                           "- Chỉ đắp sau khi kết cấu BT đạt cường độ cho phép tải và đã nghiệm thu.", "- Đắp đối xứng, không xe nặng sát tường; đá xô đạt yêu cầu kích cỡ, sạch, không lẫn bùn đất."]):
        txt(c, 25, 113 - i * 5.8, s, 9)
    # an toàn
    txt(c, 210, 88, "C. AN TOÀN LAO ĐỘNG VÀ VỆ SINH MÔI TRƯỜNG", 10, "B", "c")
    items = [("RÀO CHẮN,\nBIỂN BÁO", "Quanh hố móng, đê quai,\nlạch; cắm biển 'Cấm vào'"), ("PHÒNG HỘ\nCÁ NHÂN", "Mũ, giày, găng,\náo phản quang, dây an toàn"),
             ("ĐIỆN, MÁY\nTHI CÔNG", "Tủ điện chống giật, kiểm tra\nmáy, hiệu cẩu, trạm bơm"), ("THOÁT NƯỚC,\nHỐ LẮNG", "Bơm thoát nước hố móng,\nhố lắng trước khi xả"),
             ("RỬA XE,\nVỆ SINH", "Rửa xe ra cổng, che phủ xe chở\nđất, quét dọn đường công vụ"), ("PCCC,\nSƠ CỨU", "Bình chữa cháy, tủ thuốc,\nsố điện thoại khẩn cấp")]
    bw = 62
    for i, (a, b) in enumerate(items):
        x = 15 + i * 65.3
        rect(c, x, 62 + 0, bw, 22 + 0, lw=0.8) if False else None
    for i, (a, b) in enumerate(items):
        xx = 15 + i * 65
        rect(c, xx, 60, 62, 24, lw=0.8)
        for k, s in enumerate(a.split("\n")):
            txt(c, xx + 31, 79 - k * 4.4, s, 8.5, "B", "c")
        for k, s in enumerate(b.split("\n")):
            txt(c, xx + 31, 70 - k * 4.0, s, 7.8, "R", "c")


def main(out):
    c = canvas.Canvas(out, pagesize=(W, H))
    c.setTitle("Bản vẽ biện pháp thi công")
    c.setAuthor(PRJ["nha_thau"])
    for f in (bv01, bv02, bv03, bv04, bv05, bv06):
        f(c)
        c.showPage()
    c.save()
    print("OK", out)


if __name__ == "__main__":
    main(sys.argv[1])
