"""Xuất tiến độ thi công dạng chỉnh sửa được: Excel (Gantt tự cập nhật theo ngày) và MS Project (MSPDI .xml).
Dùng: python tools/build_tien_do_file.py <lich.xlsx> <thu_muc_ra>
"""
import datetime as dt
import sys
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).parent))
import build_tien_do as T  # noqa: E402

PD, ROWS, HOLIDAY, START, END, PRJ = T.PD, T.ROWS, T.HOLIDAY, T.START, T.END, T.PRJ
# tiền nhiệm (cùng phân đoạn): công việc -> các công việc đứng trước
PRE = {"Đào bùn, nạo vét lạch phía núi (mở mặt bằng)": ["Đắp đê quai, tạo sàn thi công, thoát nước tạm"],
       "Đào móng, vận chuyển đất đào": ["Đào bùn, nạo vét lạch phía núi (mở mặt bằng)"],
       "Ép cọc, cắt đầu cọc": ["Đào móng, vận chuyển đất đào", "Cung cấp, nghiệm thu đầu vào cọc BT300x300"],
       "Ván khuôn, đổ bê tông lót M100": ["Ép cọc, cắt đầu cọc"],
       "Cốt thép, ván khuôn, khe lún": ["Ván khuôn, đổ bê tông lót M100"],
       "Đổ bê tông M300 bền sunfat": ["Cốt thép, ván khuôn, khe lún"],
       "Đắp hoàn trả hố móng, lu nền": ["Đổ bê tông M300 bền sunfat"],
       "Đắp đá xô chân tường chắn": ["Đắp hoàn trả hố móng, lu nền"],
       "Hoàn thiện nạo vét lạch phía núi": ["Đắp đá xô chân tường chắn"]}
HC = "Hoàn thiện nạo vét lạch phía núi"
NT, BG = "Nghiệm thu hoàn thành toàn gói", "Bàn giao công trình và hồ sơ"


def build_tasks(data):
    """Danh sách phẳng: dict(level, wbs, name, pd, s, e, crit, group, preds[ids], summary)."""
    tasks, ids = [], {}
    g = -1
    gi = 0
    for name, codes in ROWS:
        if name == "#":
            g += 1
            gi += 1
            tasks.append(dict(level=1, wbs=str(gi), name=codes, group=g, summary=True, preds=[], pd=None))
            wi = 0
            continue
        wi += 1
        wbs = f"{gi}.{wi}"
        if codes[0] in ("17", "70", "71"):
            ds = [x for x in data if x[0] == codes[0]]
            tasks.append(dict(level=2, wbs=wbs, name=name, group=g, summary=False, preds=[], pd=None, key=(name, None),
                              s=min(x[2] for x in ds), e=max(x[3] for x in ds), crit=name in (NT, BG), milestone=codes[0] in ("70", "71")))
            ids[(name, None)] = len(tasks)
            continue
        tasks.append(dict(level=2, wbs=wbs, name=name, group=g, summary=True, preds=[], pd=None))
        for k, p in enumerate(PD, 1):
            ds = [x for x in data if x[0] in codes and x[1] == p]
            tasks.append(dict(level=3, wbs=f"{wbs}.{k}", name=f"{name} - {p.replace('Cọc', 'cọc')}", group=g, summary=False, pd=k,
                              s=min(x[2] for x in ds), e=max(x[3] for x in ds), crit=(k == 4 and name in T.CRIT_ROWS), key=(name, k)))
            ids[(name, k)] = len(tasks)
    for t in tasks:  # tiền nhiệm
        if t["summary"] or "key" not in t:
            continue
        nm, k = t["key"]
        if k:
            t["preds"] = [ids[(q, k)] for q in PRE.get(nm, [])]
        elif nm == NT:
            t["preds"] = [ids[(HC, k_)] for k_ in range(1, 5)]
        elif nm == BG:
            t["preds"] = [ids[(NT, None)]]
    return tasks


def ndays(s, e):
    return (e - s).days + 1 - sum(1 for h in HOLIDAY if s <= h <= e)


def write_xlsx(tasks, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tien_do_thi_cong"
    days = [START + dt.timedelta(i) for i in range((END - START).days + 1)]
    nfix = 7
    c0 = nfix + 1
    last = c0 + len(days) - 1
    thin = Side(style="thin", color="BFBFBF")
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)
    fl = lambda c: PatternFill("solid", start_color=c.lstrip("#"), end_color=c.lstrip("#"))
    F = lambda **k: Font(name="Times New Roman", **k)
    ws["A1"] = "TIẾN ĐỘ THI CÔNG"
    ws["A1"].font = F(size=18, bold=True, color="404040")
    ws["A2"] = PRJ["goi_thau"].upper()
    ws["A2"].font = F(size=12, bold=True)
    ws["A3"] = (f"Hợp đồng {PRJ['so_hop_dong']} | Chủ đầu tư: {PRJ['chu_dau_tu']} | Nhà thầu: {PRJ['nha_thau']} | "
                f"Sửa ngày Bắt đầu/Kết thúc (cột D, E): thanh tiến độ và số ngày tự cập nhật. Thanh đen và đường thẳng có mũi tên: đường găng.")
    ws["A3"].font = F(size=9.5, italic=True)
    H = 5
    heads = ["WBS", "Công việc", "Phân đoạn", "Bắt đầu", "Kết thúc", "Số ngày", "Tiền nhiệm (WBS)"]
    for i, h in enumerate(heads, 1):
        ws.merge_cells(start_row=H, start_column=i, end_row=H + 2, end_column=i)
        c = ws.cell(H, i, h)
        c.font, c.fill = F(size=10, bold=True, color="FFFFFF"), fl(T.NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(H, H + 3):
            ws.cell(r, i).border = BD
            ws.cell(r, i).fill = fl(T.NAVY)
    i = 0
    while i < len(days):
        j = i
        while j + 1 < len(days) and days[j + 1].month == days[i].month:
            j += 1
        ws.merge_cells(start_row=H, start_column=c0 + i, end_row=H, end_column=c0 + j)
        c = ws.cell(H, c0 + i, f"Tháng {days[i].month}/{days[i].year}")
        c.font, c.fill, c.alignment = F(size=10, bold=True, color="FFFFFF"), fl(T.NAVY), Alignment(horizontal="center")
        i = j + 1
    for k, d in enumerate(days):
        col = c0 + k
        a = ws.cell(H + 1, col, d)
        a.number_format = "d"
        b = ws.cell(H + 2, col, f'=CHOOSE(WEEKDAY({L(col)}{H + 1},2),"T2","T3","T4","T5","T6","T7","CN")')
        for x, sz in ((a, 8), (b, 6.5)):
            x.font, x.alignment, x.border = F(size=sz, bold=(x is a)), Alignment(horizontal="center", vertical="center"), BD
            x.fill = fl("#9A9A9A" if d in HOLIDAY else ("#D9D9D9" if d.weekday() == 6 else "#F2F2F2"))
        ws.column_dimensions[L(col)].width = 2.7
    for col, w in zip("ABCDEFG", (8, 62, 14, 12, 12, 9, 14)):
        ws.column_dimensions[col].width = w
    nl = wb.create_sheet("Ngay_le")
    nl["A1"] = "Ngày nghỉ lễ"
    for k, h in enumerate(sorted(HOLIDAY), 2):
        nl.cell(k, 1, h).number_format = "dd/mm/yyyy"
    nl.column_dimensions["A"].width = 16
    r = H + 3
    first_row = r
    row_of = {}
    for idx, t in enumerate(tasks, 1):
        row_of[idx] = r
        ws.cell(r, 1, t["wbs"])
        ws.cell(r, 2, ("      " if t["level"] == 3 else "   " if t["level"] == 2 else "") + t["name"])
        for cc in range(1, last + 1):
            ws.cell(r, cc).border = BD
        if t["summary"]:
            for cc in range(1, last + 1):
                ws.cell(r, cc).fill = fl(T.GL[t["group"]])
            ws.cell(r, 2).font = F(size=10, bold=True, color="000000")
            ws.cell(r, 1).font = F(size=9.5, bold=True)
            if t["level"] == 2:  # dòng tóm tắt: ngày = min/max của con
                kids = [i for i, x in enumerate(tasks, 1) if x["wbs"].startswith(t["wbs"] + ".")]
                a, b = row_of[kids[0]] if kids[0] in row_of else None, None
                t["_kids"] = kids
        else:
            ws.cell(r, 3, f"PĐ{t['pd']}" if t["pd"] else "Toàn tuyến").alignment = Alignment(horizontal="center")
            ws.cell(r, 4, t["s"]).number_format = "dd/mm/yyyy"
            ws.cell(r, 5, t["e"]).number_format = "dd/mm/yyyy"
            ws.cell(r, 6, f"=NETWORKDAYS.INTL(D{r},E{r},\"0000000\",Ngay_le!$A$2:$A$3)")
            ws.cell(r, 7, ", ".join(tasks[p - 1]["wbs"] for p in t["preds"]) if t["preds"] else "")
            for cc in (4, 5, 6, 7):
                ws.cell(r, cc).alignment = Alignment(horizontal="center")
            ws.cell(r, 2).font = F(size=10, bold=t.get("crit", False))
            for cc in (1, 3, 4, 5, 6, 7):
                ws.cell(r, cc).font = F(size=9.5)
            rng = f"{L(c0)}{r}:{L(last)}{r}"
            colr = "000000" if t.get("crit") else T.GC[t["group"]][1:]
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND({L(c0)}${H + 1}>=$D{r},{L(c0)}${H + 1}<=$E{r})"],
                                                          fill=PatternFill("solid", start_color=colr, end_color=colr, bgColor=colr), stopIfTrue=True))
        ws.row_dimensions[r].height = 15
        r += 1
    for t_idx, t in enumerate(tasks, 1):  # dòng tóm tắt cấp 2: ngày theo công việc con
        if t["summary"] and t["level"] == 2:
            kids = t["_kids"]
            rr, k0, k1 = row_of[t_idx], row_of[kids[0]], row_of[kids[-1]]
            ws.cell(rr, 4, f"=MIN(D{k0}:D{k1})").number_format = "dd/mm/yyyy"
            ws.cell(rr, 5, f"=MAX(E{k0}:E{k1})").number_format = "dd/mm/yyyy"
            ws.cell(rr, 6, f'=NETWORKDAYS.INTL(D{rr},E{rr},"0000000",Ngay_le!$A$2:$A$3)')
            for cc in (4, 5, 6):
                ws.cell(rr, cc).alignment, ws.cell(rr, cc).font = Alignment(horizontal="center"), F(size=9.5, bold=True)
    # ký
    r += 2
    blocks = [("ĐẠI DIỆN CHỦ ĐẦU TƯ", PRJ["chu_dau_tu"].upper(), (1, 3 + 2)),
              ("ĐẠI DIỆN TƯ VẤN QUẢN LÝ DỰ ÁN VÀ GIÁM SÁT", "CÔNG TY TNHH APAVE CHÂU Á – THÁI BÌNH DƯƠNG", (c0 + 2, c0 + 27)),
              ("ĐẠI DIỆN NHÀ THẦU THI CÔNG", PRJ["nha_thau"].upper(), (c0 + 46, c0 + 70))]
    for a, b, (x, y) in blocks:
        for k, (txt, st) in enumerate(((a, dict(size=11, bold=True, color="404040")), (b, dict(size=10, bold=True)), ("(Ký, ghi rõ họ tên, đóng dấu)", dict(size=9, italic=True)))):
            ws.merge_cells(start_row=r + k, start_column=x, end_row=r + k, end_column=y)
            c = ws.cell(r + k, x, txt)
            c.font, c.alignment = F(**st), Alignment(horizontal="center")
    ws.freeze_panes = ws.cell(H + 3, c0)
    ws.sheet_view.zoomScale = 85
    ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", 8
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"{H}:{H + 2}"
    ws.print_area = f"A1:{L(last)}{r + 9}"
    wb.save(out)
    lines = []
    for t_idx, t in enumerate(tasks, 1):
        if not t.get("crit") or t["summary"]:
            continue
        for p_ in t["preds"]:
            q = tasks[p_ - 1]
            if q.get("crit"):
                # từ cuối thanh trước (mép phải ô ngày kết thúc) đến đầu thanh sau (mép trái ô ngày bắt đầu), giữa dòng
                lines.append(((c0 - 1) + (q["e"] - START).days + 1, row_of[p_] - 1, (c0 - 1) + (t["s"] - START).days, row_of[t_idx] - 1))
    add_lines(out, lines)


def add_lines(path, lines):
    """Chèn các đường thẳng đỏ có mũi tên (Insert > Line) vào sheet 1 của file xlsx."""
    import re
    import zipfile
    anchors = []
    for k, (c1, r1, c2, r2) in enumerate(lines, 1):
        off = 95250  # nửa chiều cao dòng 15pt (EMU)
        flip = ' flipH="1"' if c2 < c1 else ""
        ca, cb = min(c1, c2), max(c1, c2)
        anchors.append(
            f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{ca}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>{off}</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{cb}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>{off}</xdr:rowOff></xdr:to>'
            f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{k + 1}" name="Duong gang {k}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>'
            f'<xdr:spPr><a:xfrm{flip}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>')
    drawing = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
               'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + "".join(anchors) + "</xdr:wsDr>")
    zin = zipfile.ZipFile(path)
    items = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    sheet = items["xl/worksheets/sheet1.xml"].decode("utf-8")
    if "xmlns:r=" not in sheet.split(">", 2)[1] + sheet.split(">", 1)[0]:
        sheet = sheet.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    # <drawing> phải đứng sau pageSetup/headerFooter và trước legacyDrawing/tableParts
    m = re.search(r"<(legacyDrawing|tableParts|extLst)\b", sheet)
    tag = '<drawing r:id="rIdDr1"/>'
    sheet = sheet[:m.start()] + tag + sheet[m.start():] if m else sheet.replace("</worksheet>", tag + "</worksheet>")
    items["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")
    rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
    rel = '<Relationship Id="rIdDr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
    if rels_name in items:
        items[rels_name] = items[rels_name].decode("utf-8").replace("</Relationships>", rel + "</Relationships>").encode("utf-8")
    else:
        items[rels_name] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rel + "</Relationships>").encode("utf-8")
    items["xl/drawings/drawing1.xml"] = drawing.encode("utf-8")
    ct = items["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>", '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in items.items():
            z.writestr(n, b)


def write_msp(tasks, out):
    iso = lambda d, hh: f"{d:%Y-%m-%d}T{hh}"
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<Project xmlns="http://schemas.microsoft.com/project">',
             f"<Name>{escape('Tien do thi cong ke chan dat tuyen 3 coc 88-126')}</Name><Title>{escape('Tiến độ thi công ' + PRJ['goi_thau'])}</Title>",
             f"<Company>{escape(PRJ['nha_thau'])}</Company>",
             "<ScheduleFromStart>1</ScheduleFromStart>", f"<StartDate>{iso(START, '08:00:00')}</StartDate>", f"<FinishDate>{iso(END, '17:00:00')}</FinishDate>",
             "<CalendarUID>1</CalendarUID>", "<DefaultStartTime>08:00:00</DefaultStartTime>", "<DefaultFinishTime>17:00:00</DefaultFinishTime>",
             "<MinutesPerDay>480</MinutesPerDay>", "<MinutesPerWeek>3360</MinutesPerWeek>", "<DaysPerMonth>30</DaysPerMonth>", "<DurationFormat>7</DurationFormat>",
             "<Calendars><Calendar><UID>1</UID><Name>7 ngay/tuan</Name><IsBaseCalendar>1</IsBaseCalendar><WeekDays>"]
    for dtp in range(1, 8):
        lines.append(f"<WeekDay><DayType>{dtp}</DayType><DayWorking>1</DayWorking><WorkingTimes>"
                     "<WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>"
                     "<WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>")
    lines.append("</WeekDays><Exceptions>")
    for k, h in enumerate(sorted(HOLIDAY), 1):
        lines.append(f"<Exception><EnteredByOccurrences>0</EnteredByOccurrences><Name>Nghi le {h:%d/%m}</Name><Type>1</Type><DayWorking>0</DayWorking>"
                     f"<TimePeriod><FromDate>{iso(h, '00:00:00')}</FromDate><ToDate>{iso(h, '23:59:00')}</ToDate></TimePeriod></Exception>")
    lines.append("</Exceptions></Calendar></Calendars><Tasks>")
    lines.append(f"<Task><UID>0</UID><ID>0</ID><Name>{escape(PRJ['goi_thau'])}</Name><Type>1</Type><IsNull>0</IsNull><OutlineNumber>0</OutlineNumber><OutlineLevel>0</OutlineLevel>"
                 f"<Start>{iso(START, '08:00:00')}</Start><Finish>{iso(END, '17:00:00')}</Finish><Summary>1</Summary></Task>")
    rng = {}  # tóm tắt: min/max của con
    for i, t in enumerate(tasks):
        if not t["summary"]:
            continue
        kids = [x for x in tasks if not x["summary"] and x["wbs"].startswith(t["wbs"] + ".") ] if True else []
        if t["level"] == 1:
            kids = [x for x in tasks if not x["summary"] and x["group"] == t["group"]]
        rng[i] = (min(x["s"] for x in kids), max(x["e"] for x in kids))
    for i, t in enumerate(tasks):
        uid = i + 1
        s, e = rng[i] if t["summary"] else (t["s"], t["e"])
        nd = ndays(s, e)
        mil = t.get("milestone")
        lines.append(f"<Task><UID>{uid}</UID><ID>{uid}</ID><Name>{escape(t['name'])}</Name><Type>0</Type><IsNull>0</IsNull><WBS>{t['wbs']}</WBS>"
                     f"<OutlineNumber>{t['wbs']}</OutlineNumber><OutlineLevel>{t['level']}</OutlineLevel><Priority>500</Priority>"
                     f"<Start>{iso(s, '08:00:00')}</Start><Finish>{iso(e, '17:00:00')}</Finish><Duration>PT{8 * nd}H0M0S</Duration><DurationFormat>7</DurationFormat>"
                     f"<Milestone>0</Milestone><Summary>{1 if t['summary'] else 0}</Summary><Critical>{1 if t.get('crit') else 0}</Critical>"
                     + ("" if t["summary"] else f"<ConstraintType>4</ConstraintType><ConstraintDate>{iso(s, '08:00:00')}</ConstraintDate>")
                     + "".join(f"<PredecessorLink><PredecessorUID>{p}</PredecessorUID><Type>1</Type><CrossProject>0</CrossProject><LinkLag>0</LinkLag><LagFormat>7</LagFormat></PredecessorLink>"
                               for p in t["preds"])
                     + "</Task>")
    lines.append("</Tasks></Project>")
    Path(out).write_text("\n".join(lines), encoding="utf-8")


def main(lich, outdir):
    data = T.load(lich)
    tasks = build_tasks(data)
    outdir = Path(outdir)
    write_xlsx(tasks, outdir / "08_Tien_do_thi_cong.xlsx")
    write_msp(tasks, outdir / "08_Tien_do_thi_cong_MSProject.xml")
    print("OK", len(tasks), "dòng")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
