"""Lập tiến độ thi công (Gantt theo ngày, A3 ngang, in đen trắng) từ lịch nghiệm thu, có đường găng.
Dùng: python tools/build_tien_do.py <lich.xlsx> <file_ra.pdf>
"""
import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).parent.parent
PRJ = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))["project"]
START, END = dt.date(2026, 8, 12), dt.date(2026, 10, 27)
HOLIDAY = {dt.date(2026, 9, 1), dt.date(2026, 9, 2)}
FD = "/usr/share/fonts/truetype/liberation/LiberationSerif-"
for n, f in (("R", "Regular"), ("B", "Bold"), ("I", "Italic")):
    pdfmetrics.registerFont(TTFont("TNR-" + n, FD + f + ".ttf"))
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]

# (tên, [mã công việc]) ; nhóm: ("#", tên nhóm)
ROWS = [("#", "A. CHUẨN BỊ MẶT BẰNG"),
        ("Đắp đê quai, tạo sàn thi công, thoát nước tạm", ["15"]),
        ("Đào bùn, nạo vét lạch phía núi (mở mặt bằng)", ["16a"]),
        ("#", "B. THI CÔNG CỌC"),
        ("Cung cấp, nghiệm thu đầu vào cọc BT300x300", ["01"]),
        ("Đào móng, vận chuyển đất đào", ["04", "05"]),
        ("Ép cọc, cắt đầu cọc", ["02", "03"]),
        ("#", "C. KẾT CẤU KÈ BÊ TÔNG"),
        ("Ván khuôn, đổ bê tông lót M100", ["10", "08"]),
        ("Cốt thép, ván khuôn, khe lún", ["13", "11", "12", "14"]),
        ("Đổ bê tông M300 bền sunfat", ["09"]),
        ("#", "D. HOÀN TRẢ, HOÀN THIỆN"),
        ("Đắp hoàn trả hố móng, lu nền", ["06"]),
        ("Đắp đá xô chân tường chắn", ["07"]),
        ("Hoàn thiện nạo vét lạch phía núi", ["16b"]),
        ("#", "E. CHUNG TOÀN TUYẾN"),
        ("Biện pháp thi công trọn gói", ["17"]),
        ("Nghiệm thu hoàn thành toàn gói", ["70"]),
        ("Bàn giao công trình và hồ sơ", ["71"])]
# đường găng: các công việc ở phân đoạn 4 (cọc 118–126) rồi các mốc toàn gói
CRIT = [("Đắp đê quai, tạo sàn thi công, thoát nước tạm", "Đào bùn, nạo vét lạch phía núi (mở mặt bằng)"),
        ("Đào bùn, nạo vét lạch phía núi (mở mặt bằng)", "Đào móng, vận chuyển đất đào"),
        ("Đào móng, vận chuyển đất đào", "Ép cọc, cắt đầu cọc"),
        ("Cung cấp, nghiệm thu đầu vào cọc BT300x300", "Ép cọc, cắt đầu cọc"),
        ("Ép cọc, cắt đầu cọc", "Ván khuôn, đổ bê tông lót M100"),
        ("Ván khuôn, đổ bê tông lót M100", "Cốt thép, ván khuôn, khe lún"),
        ("Cốt thép, ván khuôn, khe lún", "Đổ bê tông M300 bền sunfat"),
        ("Đổ bê tông M300 bền sunfat", "Đắp hoàn trả hố móng, lu nền"),
        ("Đắp hoàn trả hố móng, lu nền", "Đắp đá xô chân tường chắn"),
        ("Đắp đá xô chân tường chắn", "Hoàn thiện nạo vét lạch phía núi"),
        ("Hoàn thiện nạo vét lạch phía núi", "Nghiệm thu hoàn thành toàn gói"),
        ("Nghiệm thu hoàn thành toàn gói", "Bàn giao công trình và hồ sơ")]
CRIT_ROWS = {a for a, _ in CRIT} | {b for _, b in CRIT}


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Chi tiet nghiem thu"]
    out = []
    for r in ws.iter_rows(min_row=9, values_only=True):
        if r[0] is None or r[6] is None:
            continue
        muc = str(r[1] or "")
        if not muc:
            muc = "70" if str(r[2]).startswith("Nghiệm thu") else "71"
        elif muc == "16":
            muc = "16b" if str(r[2]).lower().startswith("hoàn thiện") else "16a"
        out.append((muc, r[3], r[6].date(), r[7].date()))
    return out


def main(lich, out):
    data = load(lich)
    W, H = landscape(A3)
    mm = 72 / 25.4
    c = canvas.Canvas(out, pagesize=(W, H))
    c.setTitle("Tiến độ thi công"), c.setAuthor(PRJ["nha_thau"])
    X0, X1, YT = 14 * mm, W - 14 * mm, H - 40 * mm
    cols = [("STT", 11), ("Công việc chính", 88), ("Bắt đầu", 21), ("Kết thúc", 21), ("Số ngày", 15)]
    xs = [X0]
    for _, w in cols:
        xs.append(xs[-1] + w * mm)
    GX = xs[-1]
    days = [START + dt.timedelta(i) for i in range((END - START).days + 1)]
    dw = (X1 - GX) / len(days)
    dx = lambda i: GX + i * dw  # i: chỉ số ngày (0 = 12/08)
    ix = lambda d: (d - START).days

    # khung trang
    c.setLineWidth(1.2)
    c.rect(8 * mm, 8 * mm, W - 16 * mm, H - 16 * mm)
    c.setLineWidth(0.4)
    c.rect(9.5 * mm, 9.5 * mm, W - 19 * mm, H - 19 * mm)
    # tiêu đề
    c.setFont("TNR-B", 22)
    c.drawCentredString(W / 2, H - 22 * mm, "TIẾN ĐỘ THI CÔNG")
    c.setFont("TNR-B", 12.5)
    c.drawCentredString(W / 2, H - 29 * mm, PRJ["goi_thau"].upper())
    c.setFont("TNR-I", 9.5)
    c.drawCentredString(W / 2, H - 34.5 * mm, f"Hợp đồng {PRJ['so_hop_dong']}  |  Địa điểm: {PRJ['dia_diem']}  |  Thời gian: {START:%d/%m/%Y} – {END:%d/%m/%Y}  |  "
                        "Thanh đen và mũi tên: đường găng  |  Số trong thanh: phân đoạn 1–4 (cọc 88–98, 98–108, 108–118, 118–126)")

    # tiêu đề cột
    hh = (6.5, 5, 4.5)
    HT = sum(hh) * mm
    c.setFillGray(0.82)
    c.rect(X0, YT - HT, X1 - X0, HT, fill=1, stroke=0)
    c.setFillGray(0)
    c.setLineWidth(0.6)
    for i, (t, _) in enumerate(cols):
        c.rect(xs[i], YT - HT, xs[i + 1] - xs[i], HT, fill=0)
        c.setFont("TNR-B", 10)
        c.drawCentredString((xs[i] + xs[i + 1]) / 2, YT - HT / 2 - 1.2 * mm, t)
    k = 0
    while k < len(days):
        j = k
        while j + 1 < len(days) and days[j + 1].month == days[k].month:
            j += 1
        c.rect(dx(k), YT - hh[0] * mm, dx(j + 1) - dx(k), hh[0] * mm)
        c.setFont("TNR-B", 9.5)
        c.drawCentredString((dx(k) + dx(j + 1)) / 2, YT - hh[0] * mm + 1.8 * mm, f"Tháng {days[k].month}/{days[k].year}")
        k = j + 1
    for i, d in enumerate(days):
        y1, y2 = YT - hh[0] * mm, YT - (hh[0] + hh[1]) * mm
        c.setFillGray(0.55 if d in HOLIDAY else (0.9 if d.weekday() == 6 else 1))
        c.rect(dx(i), y2, dw, hh[1] * mm, fill=1)
        c.rect(dx(i), y2 - hh[2] * mm, dw, hh[2] * mm, fill=1)
        c.setFillGray(0)
        c.setFont("TNR-B", 6.3)
        c.drawCentredString(dx(i) + dw / 2, y2 + 1.4 * mm, str(d.day))
        c.setFont("TNR-R", 5.2)
        c.drawCentredString(dx(i) + dw / 2, y2 - 3.1 * mm, ["T2", "T3", "T4", "T5", "T6", "T7", "CN"][d.weekday()])

    # dòng
    RH, GH = 8.6 * mm, 6.2 * mm
    y = YT - HT
    pos = {}  # tên -> (y_top, segs)
    stt = 0
    body_top = y
    for name, codes in ROWS:
        if name == "#":
            y -= GH
            c.setFillGray(0.88)
            c.rect(X0, y, X1 - X0, GH, fill=1, stroke=1)
            c.setFillGray(0)
            c.setFont("TNR-B", 10.5)
            c.drawString(xs[1] + 2 * mm, y + 2.1 * mm, codes)
            continue
        y -= RH
        stt += 1
        segs = []
        if codes[0] in ("17", "70", "71"):
            ds = [x for x in data if x[0] == codes[0]]
            segs = [(None, min(x[2] for x in ds), max(x[3] for x in ds))]
        else:
            for p in PD:
                ds = [x for x in data if x[0] in codes and x[1] == p]
                if ds:
                    segs.append((PD.index(p) + 1, min(x[2] for x in ds), max(x[3] for x in ds)))
        s0, e0 = min(s[1] for s in segs), max(s[2] for s in segs)
        crit = name in CRIT_ROWS
        pos[name] = (y, segs)
        # nền hàng + cột bảng
        for i in range(len(cols)):
            c.rect(xs[i], y, xs[i + 1] - xs[i], RH, fill=0)
        c.setFont("TNR-R", 10)
        c.drawCentredString((xs[0] + xs[1]) / 2, y + RH / 2 - 1.3 * mm, str(stt))
        c.setFont("TNR-B" if crit else "TNR-R", 9.6)
        words, line1, lines = name.split(), "", []
        for w_ in words:
            if pdfmetrics.stringWidth(line1 + " " + w_, "TNR-R", 9.6) > (cols[1][1] - 4) * mm:
                lines.append(line1.strip())
                line1 = ""
            line1 += " " + w_
        lines.append(line1.strip())
        for li, t in enumerate(lines):
            c.drawString(xs[1] + 2 * mm, y + RH / 2 + (len(lines) - 1) * 2.1 * mm - li * 4.2 * mm - 1.2 * mm, t)
        c.setFont("TNR-R", 9.5)
        for i, t in ((2, f"{s0:%d/%m/%Y}"), (3, f"{e0:%d/%m/%Y}"), (4, str((e0 - s0).days + 1 - sum(1 for h in HOLIDAY if s0 <= h <= e0)))):
            c.drawCentredString((xs[i] + xs[i + 1]) / 2, y + RH / 2 - 1.3 * mm, t)
        # lưới ngày
        for i, d in enumerate(days):
            c.setLineWidth(0.15 if d.weekday() else 0.5)
            c.setStrokeGray(0.75 if d.weekday() else 0.45)
            c.line(dx(i), y, dx(i), y + RH)
            if d in HOLIDAY:
                c.setFillGray(0.86)
                c.rect(dx(i), y, dw, RH, fill=1, stroke=0)
        c.setStrokeGray(0)
        c.setLineWidth(0.6)
        c.line(GX, y, X1, y)
        c.line(X1, y, X1, y + RH)
        # thanh
        bh = 4.8 * mm
        for pdn, a, b in segs:
            x_a, x_b = dx(ix(a)), dx(ix(b) + 1)
            cy = y + RH / 2
            if codes[0] in ("70", "71"):
                cx, r_ = x_b - dw / 2, 3.1 * mm
                c.setFillGray(0)
                p = c.beginPath()
                p.moveTo(cx, cy + r_), p.lineTo(cx + r_, cy), p.lineTo(cx, cy - r_), p.lineTo(cx - r_, cy), p.close()
                c.drawPath(p, fill=1, stroke=1)
                continue
            if codes[0] == "17":
                c.setFillGray(0.55)
                c.rect(x_a, cy - 1.4 * mm, x_b - x_a, 2.8 * mm, fill=1, stroke=0)
                c.setStrokeGray(0), c.setLineWidth(0.5)
                c.rect(x_a, cy - 1.4 * mm, x_b - x_a, 2.8 * mm, fill=0)
                continue
            is_c = pdn == 4 and crit
            c.setLineWidth(0.5 if not is_c else 1.0)
            c.setFillGray(0 if is_c else 0.8)
            c.rect(x_a, cy - bh / 2, x_b - x_a, bh, fill=1, stroke=1)
            c.setFillGray(1 if is_c else 0)
            c.setFont("TNR-B", 6.4)
            c.drawCentredString((x_a + x_b) / 2, cy - 0.9 * mm, str(pdn))
        c.setFillGray(0)
    # khung ngoài bảng
    c.setLineWidth(1.2)
    c.rect(X0, y, X1 - X0, YT - y, fill=0)
    c.setLineWidth(1.0)
    c.line(GX, YT, GX, y)

    # đường găng: liên kết kết thúc - bắt đầu (kiểu MS Project): ra từ cuối thanh, vào đầu thanh kế tiếp từ bên trái
    def bar_ends(n):
        yy, sg = pos[n]
        s_, e_ = sg[-1][1], sg[-1][2]
        xa, xb = dx(ix(e_) + 1), dx(ix(s_))
        if n in ("Nghiệm thu hoàn thành toàn gói", "Bàn giao công trình và hồ sơ"):
            cx = dx(ix(e_) + 1) - dw / 2
            xa, xb = cx + 3.1 * mm, cx - 3.1 * mm
        return yy + RH / 2, xa, xb

    c.setStrokeGray(0), c.setFillGray(0), c.setLineWidth(1.1)
    for a_, b_ in CRIT:
        ya, xa, _ = bar_ends(a_)
        yb, _, xb = bar_ends(b_)
        xe, xs = xa + 1.6 * mm, xb - 1.6 * mm
        if a_.startswith("Nghiệm thu hoàn thành"):  # mốc nối mốc: đường thẳng đứng
            cx = xa - 3.1 * mm
            c.line(cx, ya - 3.1 * mm, cx, yb + 3.1 * mm + 0.2 * mm)
            ah = c.beginPath()
            ah.moveTo(cx, yb + 3.1 * mm), ah.lineTo(cx - 1.0 * mm, yb + 5.1 * mm), ah.lineTo(cx + 1.0 * mm, yb + 5.1 * mm), ah.close()
            c.drawPath(ah, fill=1, stroke=0)
            continue
        if xs >= xe:
            pts = [(xa, ya), (xs, ya), (xs, yb), (xb, yb)]
        else:
            ym = ya - RH / 2
            pts = [(xa, ya), (xe, ya), (xe, ym), (xs, ym), (xs, yb), (xb, yb)]
        p = c.beginPath()
        p.moveTo(*pts[0])
        for q in pts[1:-1]:
            p.lineTo(*q)
        p.lineTo(pts[-1][0] - 1.6 * mm, pts[-1][1])
        c.setLineJoin(1)
        c.drawPath(p, fill=0, stroke=1)
        ah = c.beginPath()
        ah.moveTo(xb, yb), ah.lineTo(xb - 2.0 * mm, yb + 1.0 * mm), ah.lineTo(xb - 2.0 * mm, yb - 1.0 * mm), ah.close()
        c.drawPath(ah, fill=1, stroke=0)

    # khối ký, đóng dấu (3 khối cân đối)
    blocks = [("ĐẠI DIỆN CHỦ ĐẦU TƯ", PRJ["chu_dau_tu"].upper()),
              ("ĐẠI DIỆN TƯ VẤN QUẢN LÝ DỰ ÁN", "CÔNG TY TNHH APAVE CHÂU Á – THÁI BÌNH DƯƠNG"),
              ("ĐẠI DIỆN NHÀ THẦU THI CÔNG", PRJ["nha_thau"].upper())]
    yy = y - 12 * mm
    for i, (a, b) in enumerate(blocks):
        cx = X0 + (X1 - X0) * (i * 2 + 1) / 6
        c.setFillGray(0)
        c.setFont("TNR-B", 12)
        c.drawCentredString(cx, yy, a)
        c.setFont("TNR-B", 10.5)
        c.drawCentredString(cx, yy - 5.5 * mm, b)
        c.setFont("TNR-I", 10)
        c.drawCentredString(cx, yy - 11 * mm, "(Ký, ghi rõ họ tên, đóng dấu)")
    c.showPage()
    c.save()
    print("OK", out, "y_cuoi_bang_mm", round(y / mm, 1))


main(sys.argv[1], sys.argv[2])
