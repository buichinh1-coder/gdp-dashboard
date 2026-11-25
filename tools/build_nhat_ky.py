"""Lập Nhật ký thi công xây dựng công trình (mẫu TCVN 4055, phần nhà thầu) từ form mẫu và lịch nghiệm thu.

Mô tả công việc từng ngày lấy theo ngày bắt đầu - hoàn thành của lịch (đã được xác nhận là ngày thực tế) và các ngày nghiệm thu
trong records.json. Điều kiện thi công (thời tiết, nhân lực, máy), khối lượng và chữ ký đội trưởng để trống cho nhà thầu ghi thực tế.
Dùng: python tools/build_nhat_ky.py <lich.xlsx> <file_ra.docx>
"""
import copy
import datetime as dt
import json
import sys
from pathlib import Path

import docx
import openpyxl
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
import qaqc_docx as Q  # noqa: E402

DB = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))
PRJ, RECS = DB["project"], DB["records"]
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]
START, END = dt.date(2026, 8, 12), dt.date(2026, 10, 27)  # theo hợp đồng 12.08/2026/HĐKT/ĐGC-GL
HOLIDAY = {dt.date(2026, 9, 1), dt.date(2026, 9, 2)}
WD = ["Thứ hai", "Thứ ba", "Thứ tư", "Thứ năm", "Thứ sáu", "Thứ bảy", "Chủ nhật"]
CV_NAME = {"01": "cung cấp cọc BT300x300", "02": "ép cọc", "03": "cắt, đập đầu cọc", "04": "đào móng bằng máy", "05": "vận chuyển đất đào",
           "06": "đắp hoàn trả hố móng, lu nền (độ chặt K95)", "07": "đắp đá xô chân tường chắn", "08": "đổ bê tông lót M100",
           "09": "đổ bê tông M300 (móng, thân, khóa đỉnh)", "10": "ván khuôn bê tông lót", "11": "ván khuôn cấu kiện",
           "12": "ván khuôn khóa đỉnh", "13": "gia công, lắp dựng cốt thép", "14": "thi công khe lún",
           "15": "đắp đê quai, tạo sàn thi công (độ chặt K95)", "16": "đào bùn, nạo vét lạch phía núi"}
SHORT = {"15": "Đắp đê quai", "16": "Đào bùn, nạo vét", "02+03": "Nghiệm thu cọc (ép cọc, cắt đầu cọc)", "04+05": "Đào móng, vận chuyển đất đào",
         "10+08": "Ván khuôn, đổ bê tông lót", "13+11+12+14": "Cốt thép, ván khuôn, khe lún trước đổ bê tông", "09": "Đổ bê tông M300",
         "06": "Đắp hoàn trả hố móng", "07": "Đắp đá xô chân tường chắn"}
dmy = lambda d: d.strftime("%d/%m/%Y")
num = lambda r: r["so"].split()[-1]


def load_rows(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Chi tiet nghiem thu"]
    rows = [[ws.cell(r, c).value for c in range(1, 15)] for r in range(1, ws.max_row + 1)]
    return [v for v in rows if str(v[0]).isdigit()]


def set_cell(cell, lines, size=10.5, bold=False, align=None):
    """Ghi các dòng vào ô, mỗi dòng một đoạn, cỡ chữ cố định."""
    ps = cell.paragraphs
    for p in ps[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    first = True
    for ln in lines:
        p = p0 if first else cell.add_paragraph()
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(ln)
        run.font.size, run.bold, run.font.name = Pt(size), bold, "Times New Roman"
        first = False


def day_lines(d, rows):
    out = []
    if d in HOLIDAY:
        out.append("Nghỉ lễ Quốc khánh (theo hợp đồng).")
        return out
    by = {}
    for v in rows:
        muc = str(v[1] or "")
        if muc == "17" or v[6] is None:
            continue
        s, e = v[6].date(), v[7].date()
        if not (s <= d <= e):
            continue
        if not muc:
            name = "hoàn thiện hồ sơ chất lượng, hoàn công, bàn giao công trình"
        elif muc == "16" and str(v[2]).lstrip().lower().startswith("hoàn thiện"):
            name = "hoàn thiện nạo vét lạch phía núi (30% khối lượng)"
        else:
            name = CV_NAME[muc]
        by.setdefault(v[3], [])
        if name not in by[v[3]]:
            by[v[3]].append(name)
    order = sorted(by, key=lambda k: PD.index(k) if k in PD else 9)
    for vt in order:
        txt = "; ".join(by[vt])
        out.append(f"- {vt}: {txt[0].upper() + txt[1:]}." if vt in PD else f"- {txt[0].upper() + txt[1:]}.")
    iso = d.isoformat()
    for r in RECS:
        if r["type"] == "CV" and r.get("lay_mau") and r.get("ngay_lay_mau") and r["ngay_lay_mau"] <= iso <= (r.get("ngay_lay_mau_den") or r["ngay_lay_mau"]):
            out.append(f"- Lấy mẫu bê tông M300 {r['vi_tri'].lower()} (biên bản lấy mẫu số {r['so_lm'].split()[-1]}/LM).")
    for r in RECS:
        if r["ngay"] != iso:
            continue
        if r["type"] == "VL":
            out.append(f"- Nghiệm thu vật liệu đầu vào số {num(r)}/CV/VL-TB: {r['hang_muc'].split(' (')[0].split(',')[0]}.")
        elif r["type"] == "CV":
            out.append(f"- Nghiệm thu công việc số {num(r)}/CV: {SHORT[r['muc_hd']]} ({r['vi_tri'].lower()}).")
        elif r["type"] == "GD":
            out.append(f"- Nghiệm thu giai đoạn thi công cọc toàn tuyến số {num(r)}/CV/GĐ.")
    if not out:
        out.append("- Không có công việc theo tiến độ.")
    return out


def main(lich, out_path):
    rows = load_rows(lich)
    tpl = docx.Document(str(Q.TEMPLATE))
    body = tpl.element.body
    src = list(body)
    sect = src[-1]
    for el in src[:-1]:
        body.remove(el)
    s = tpl.sections[-1]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Cm(1.5), Cm(1.5), Cm(2.5), Cm(1.5)
    els = {}
    for i in range(837, 901):  # từ tên bìa đến Bảng 5
        e = copy.deepcopy(src[i])
        sect.addprevious(e)
        els[i] = e
    P = lambda i: els[i]
    T = lambda i: Table(els[i], tpl)

    # ngắt trang: bìa | trang 1 | trang 2 | bảng
    for i in (843, 868):
        ppr = P(i).find(qn("w:pPr"))
        if ppr is None:
            ppr = P(i).makeelement(qn("w:pPr"), {})
            P(i).insert(0, ppr)
        ppr.append(ppr.makeelement(qn("w:pageBreakBefore"), {}))
    for i in (842, 867, 840):  # nhãn chú thích của form mẫu
        Q._set_text(P(i), "")

    # --- bìa
    Q._set_text(P(839), f"Nhà thầu thi công xây dựng: {PRJ['nha_thau'].upper()}")
    Q._set_text(P(841), f"Công trình: {PRJ['goi_thau']}")
    # --- trang 1
    Q._set_text(P(844), "Số: 01")
    Q._set_text(P(847), PRJ["goi_thau"])
    Q._set_text(P(848), f"- Địa chỉ công trình: {PRJ['dia_diem']}")
    Q._set_text(P(849), f"- Cơ quan giao thầu: {PRJ['chu_dau_tu']}")
    Q._set_text(P(853), f"Theo tiến độ nêu trong hợp đồng: bắt đầu {dmy(START)}, kết thúc {dmy(END)}")
    Q._set_text(P(854), f"Theo thực tế: bắt đầu {dmy(START)}, kết thúc {dmy(END)}")
    # --- trang 2
    Q._set_text(P(875), f"- Họ tên, chữ ký người phụ trách thi công công trình (hạng mục công trình) và quản lý quyển nhật ký: "
                        f"{PRJ['chi_huy_truong']} - Chỉ huy trưởng công trình")

    # --- Bảng 1: cán bộ kỹ thuật nhà thầu
    t1 = T(884)
    set_cell(t1.rows[1].cells[0], ["1"], 12, align=1)
    set_cell(t1.rows[1].cells[1], [PRJ["chi_huy_truong"]], 12)
    set_cell(t1.rows[1].cells[3], ["Chỉ huy trưởng công trình"], 12)
    set_cell(t1.rows[1].cells[4], [dmy(START)], 12, align=1)
    set_cell(t1.rows[1].cells[5], [dmy(END)], 12, align=1)

    # --- Bảng 2: bản kê biên bản nghiệm thu
    evs = sorted(RECS, key=lambda r: (r["ngay"], {"VL": 0, "CV": 1, "GD": 2}[r["type"]], int(num(r))))
    t2 = T(888)
    tmpl_tr = copy.deepcopy(t2.rows[1]._tr)
    for _ in range(len(evs) - 1):
        t2._tbl.append(copy.deepcopy(tmpl_tr))
    for k, r in enumerate(evs):
        row = t2.rows[k + 1]
        if r["type"] == "VL":
            nm = f"Biên bản nghiệm thu vật liệu đầu vào số {num(r)}/CV/VL-TB: {r['hang_muc'].split(' (')[0]}"
        elif r["type"] == "CV":
            nm = f"Biên bản nghiệm thu công việc số {num(r)}/CV: {SHORT[r['muc_hd']]} - {r['vi_tri']}"
        else:
            nm = f"Biên bản nghiệm thu giai đoạn số {num(r)}/CV/GĐ: thi công cọc toàn tuyến"
        set_cell(row.cells[0], [str(k + 1)], 11, align=1)
        set_cell(row.cells[1], [nm], 11, align=0)
        set_cell(row.cells[2], [dmy(dt.date.fromisoformat(r["ngay"]))], 11, align=1)

    # --- Bảng 4: tình hình thi công hàng ngày
    note = copy.deepcopy(P(894))
    P(894).addnext(note)
    Q._set_text(note, "Mô tả công việc lập theo tiến độ thi công trong hợp đồng và các ngày nghiệm thu đã xác nhận. Cột điều kiện thi công "
                      "(thời tiết, nhân lực, máy móc), khối lượng và chữ ký đội trưởng do nhà thầu ghi theo thực tế.")
    for r_ in note.iter(qn("w:r")):
        rpr = r_.find(qn("w:rPr"))
        if rpr is not None:
            for tg in ("w:sz", "w:szCs"):
                x = rpr.find(qn(tg))
                if x is not None:
                    x.set(qn("w:val"), "22")
    t4 = T(896)
    tmpl_tr = copy.deepcopy(t4.rows[1]._tr)
    n_days = (END - START).days + 1
    for _ in range(n_days - 1):
        t4._tbl.append(copy.deepcopy(tmpl_tr))
    for k in range(n_days):
        d = START + dt.timedelta(k)
        row = t4.rows[k + 1]
        set_cell(row.cells[0], [dmy(d), WD[d.weekday()]], 10.5, align=1)
        set_cell(row.cells[1], day_lines(d, rows), 10.5, align=0)
    for p in tpl.paragraphs:  # chữ trong các đoạn (ngoài bảng) tối thiểu 13pt như các biểu mẫu khác
        pass
    for el in els.values():
        if el.tag == qn("w:p"):
            Q._chuan_font(el)
    tpl.save(out_path)
    print("OK", n_days, "ngày;", len(evs), "biên bản trong bản kê")


main(sys.argv[1], sys.argv[2])
