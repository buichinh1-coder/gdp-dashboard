"""Lập tiến độ thi công (biểu đồ Gantt theo ngày) từ lịch nghiệm thu.
Dùng: python tools/build_tien_do.py <lich.xlsx> <file_ra.xlsx>
"""
import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ROOT = Path(__file__).parent.parent
PRJ = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))["project"]
START, END = dt.date(2026, 8, 12), dt.date(2026, 10, 27)
HOLIDAY = {dt.date(2026, 9, 1), dt.date(2026, 9, 2)}
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]
NAME = {"01": "Cung cấp, nghiệm thu đầu vào cọc BT300x300 M300 bền sunfat R7", "02": "Ép cọc BT300x300", "03": "Cắt, đập đầu cọc",
        "04": "Đào móng bằng máy (đất cấp II)", "05": "Vận chuyển đất đào", "06": "Đắp hoàn trả hố móng, lu nền (K95)",
        "07": "Đắp đá xô chân tường chắn", "08": "Đổ bê tông lót M100", "09": "Đổ bê tông M300 bền sunfat (móng, thân, khóa đỉnh)",
        "10": "Ván khuôn bê tông lót", "11": "Ván khuôn cấu kiện móng, thân", "12": "Ván khuôn khóa đỉnh", "13": "Gia công, lắp dựng cốt thép",
        "14": "Thi công khe lún, chi tiết thoát nước", "15": "Đắp đê quai, tạo sàn thi công, thoát nước tạm (K95)",
        "16": "Đào bùn, nạo vét lạch phía núi (đợt mở mặt bằng)", "17": "Biện pháp thi công trọn gói (duy trì đường công vụ, thoát nước...)"}
COL = {"01": "7F7F7F", "02": "1F4E79", "03": "2E75B6", "04": "C55A11", "05": "ED7D31", "06": "7F6000", "07": "BF9000", "08": "A9A9A9",
       "09": "C00000", "10": "548235", "11": "70AD47", "12": "548235", "13": "7030A0", "14": "00808A", "15": "9C6500", "16": "8E6F3E", "17": "595959"}
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
F = lambda **k: Font(name="Times New Roman", **k)


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Chi tiet nghiem thu"]
    out = []
    for r in ws.iter_rows(min_row=9, values_only=True):
        if r[0] is None or r[6] is None:
            continue
        muc = str(r[1] or "")
        ten = NAME.get(muc) or str(r[2])
        if muc == "16" and str(r[2]).lower().startswith("hoàn thiện"):
            ten = "Hoàn thiện nạo vét lạch phía núi (30% khối lượng còn lại)"
        out.append(dict(muc=muc, ten=ten, vt=r[3], dvt=r[4], kl=r[5], s=r[6].date(), e=r[7].date(), nt=r[8].date() if r[8] else None))
    return out


def main(lich, out):
    rows = load(lich)
    days = [START + dt.timedelta(i) for i in range((END - START).days + 1)]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tien_do_thi_cong"
    nfix = 8  # STT, công việc, ĐVT, KL, bắt đầu, kết thúc, số ngày, ghi chú
    c0 = nfix + 1
    last = c0 + len(days) - 1
    ws["A1"] = "TIẾN ĐỘ THI CÔNG"
    ws["A2"] = PRJ["goi_thau"]
    ws["A3"] = (f"Hợp đồng {PRJ['so_hop_dong']} | Chủ đầu tư: {PRJ['chu_dau_tu']} | Nhà thầu: {PRJ['nha_thau']} | "
                f"Thời gian: {START:%d/%m/%Y} – {END:%d/%m/%Y} (làm 7 ngày/tuần, nghỉ lễ 01–02/09)")
    ws["A1"].font, ws["A2"].font, ws["A3"].font = F(size=16, bold=True), F(size=12, bold=True), F(size=10, italic=True)
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
        ws.cell(r, 1).alignment = Alignment(horizontal="center")
    H = 5  # hàng tiêu đề (tháng ở H-1... dùng 3 hàng: tháng, ngày, thứ)
    heads = ["STT", "Công việc", "ĐVT", "Khối lượng", "Bắt đầu", "Kết thúc", "Số ngày", "Nghiệm thu (KH)"]
    for i, h in enumerate(heads, 1):
        ws.merge_cells(start_row=H, start_column=i, end_row=H + 2, end_column=i)
        c = ws.cell(H, i, h)
        c.font, c.alignment, c.fill = F(size=10, bold=True), Alignment(horizontal="center", vertical="center", wrap_text=True), fill("D9E1F2")
        for r in range(H, H + 3):
            ws.cell(r, i).border = BD
    # hàng tháng (gộp theo tháng)
    i = 0
    while i < len(days):
        j = i
        while j + 1 < len(days) and days[j + 1].month == days[i].month:
            j += 1
        ws.merge_cells(start_row=H, start_column=c0 + i, end_row=H, end_column=c0 + j)
        c = ws.cell(H, c0 + i, f"Tháng {days[i].month}/{days[i].year}")
        c.font, c.alignment, c.fill = F(size=10, bold=True), Alignment(horizontal="center"), fill("D9E1F2")
        for k in range(i, j + 1):
            ws.cell(H, c0 + k).border = BD
        i = j + 1
    for k, d in enumerate(days):
        for r, txt, sz in ((H + 1, d.day, 8), (H + 2, ["T2", "T3", "T4", "T5", "T6", "T7", "CN"][d.weekday()], 6.5)):
            c = ws.cell(r, c0 + k, txt)
            c.font, c.alignment, c.border = F(size=sz, bold=(r == H + 1)), Alignment(horizontal="center", vertical="center"), BD
            c.fill = fill("F4B183" if d in HOLIDAY else ("EDEDED" if d.weekday() == 6 else "FFFFFF"))
        ws.column_dimensions[L(c0 + k)].width = 2.6
    for col, w in zip("ABCDEFGH", (5, 58, 7, 11, 11, 11, 8, 12)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[H + 2].height = 12

    r = H + 3
    stt = 0

    def group(title):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nfix)
        c = ws.cell(r, 1, title)
        c.font, c.fill = F(size=10.5, bold=True), fill("FFF2CC")
        for cc in range(1, last + 1):
            ws.cell(r, cc).border = BD
            if cc >= c0:
                ws.cell(r, cc).fill = fill("FFF2CC")
        r += 1

    def line(v):
        nonlocal r, stt
        stt += 1
        n = (v["e"] - v["s"]).days + 1 - sum(1 for h in HOLIDAY if v["s"] <= h <= v["e"])
        vals = [stt, v["ten"], v["dvt"], v["kl"], v["s"], v["e"], n, v["nt"]]
        for i, x in enumerate(vals, 1):
            c = ws.cell(r, i, x)
            c.font, c.border = F(size=9.5), BD
            c.alignment = Alignment(horizontal="left" if i == 2 else "center", vertical="center", wrap_text=(i == 2))
            if i in (5, 6, 8) and x:
                c.number_format = "dd/mm/yyyy"
            if i == 4 and x is not None:
                c.number_format = "#,##0.00"
        for k, d in enumerate(days):
            c = ws.cell(r, c0 + k)
            c.border = BD
            if d in HOLIDAY:
                c.fill = fill("FBE5D6")
            if v["s"] <= d <= v["e"] and d not in HOLIDAY:
                c.fill = fill(COL.get(v["muc"], "2E75B6"))
            if v["nt"] and d == v["nt"]:
                c.value, c.font, c.alignment = "◆", F(size=8, bold=True, color="FFFFFF" if v["s"] <= d <= v["e"] else "000000"), Alignment(horizontal="center")
        ws.row_dimensions[r].height = 24 if len(v["ten"]) > 55 else 15
        r += 1

    for p in PD:
        group(f"PHÂN ĐOẠN {PD.index(p) + 1}: {p.replace('Cọc ', 'cọc ').replace('-', ' – ')}")
        for v in sorted([x for x in rows if x["vt"] == p], key=lambda x: (x["s"], x["e"], x["muc"])):
            line(v)
    group("TOÀN TUYẾN: cọc 88 – 126")
    for v in sorted([x for x in rows if x["vt"] not in PD], key=lambda x: (x["s"], x["e"])):
        line(v)
    r += 1
    notes = ["Ghi chú:",
             "- Thanh màu: thời gian thi công theo lịch nghiệm thu (bắt đầu – hoàn thành); ◆: ngày nghiệm thu theo kế hoạch; ô cam: nghỉ lễ 01–02/09 theo hợp đồng.",
             "- Khối lượng từng phân đoạn là số phân bổ tạm trong lịch nghiệm thu, cần cập nhật theo bản vẽ triển khai.",
             "- Các công việc bê tông M300 có lấy mẫu R7/R28 theo từng đợt đổ; thời gian bảo dưỡng và thí nghiệm R28 nằm ngoài khung thi công trên biểu đồ.",
             "- Tiến độ lập theo ngày kế hoạch đã được xác nhận là ngày thực tế; chưa kèm nhân lực và thiết bị (không có dữ liệu)."]
    for t in notes:
        ws.cell(r, 1, t).font = F(size=9.5, bold=(t == "Ghi chú:"), italic=(t != "Ghi chú:"))
        r += 1
    ws.freeze_panes = ws.cell(H + 3, c0)
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", 8  # A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = 0.4, 0.4, 0.5, 0.5
    ws.print_title_rows = f"{H}:{H + 2}"
    ws.print_area = f"A1:{L(last)}{r}"

    # --- bảng mốc chính
    m = wb.create_sheet("Moc_chinh")
    m["A1"] = "CÁC MỐC TIẾN ĐỘ CHÍNH"
    m["A1"].font = F(size=14, bold=True)
    for i, h in enumerate(["Mốc", "Ngày"], 1):
        c = m.cell(3, i, h)
        c.font, c.fill, c.border = F(bold=True), fill("D9E1F2"), BD
    first = lambda muc, p: min(x["s"] for x in rows if x["muc"] == muc and x["vt"] == p)
    lastd = lambda muc, p: max(x["e"] for x in rows if x["muc"] == muc and x["vt"] == p)
    ms = [("Khởi công theo hợp đồng", START)]
    for p in PD:
        ms.append((f"{p}: hoàn thành ép cọc", lastd("02", p)))
    for p in PD:
        ms.append((f"{p}: đổ bê tông M300", f"{first('09', p):%d/%m/%Y} – {lastd('09', p):%d/%m/%Y}"))
    for p in PD:
        ms.append((f"{p}: hoàn thành đắp đá xô chân tường chắn", lastd("07", p)))
    ms += [("Nghiệm thu hoàn thành toàn gói", dt.date(2026, 10, 26)), ("Bàn giao công trình và hồ sơ; kết thúc hợp đồng", END)]
    for k, (a, b) in enumerate(ms, 4):
        m.cell(k, 1, a).font = F()
        c = m.cell(k, 2, b)
        c.font, c.alignment = F(), Alignment(horizontal="center")
        if isinstance(b, dt.date):
            c.number_format = "dd/mm/yyyy"
        m.cell(k, 1).border = m.cell(k, 2).border = BD
    m.column_dimensions["A"].width, m.column_dimensions["B"].width = 55, 25
    wb.save(out)
    print("OK", len(rows), "công việc")


main(sys.argv[1], sys.argv[2])
