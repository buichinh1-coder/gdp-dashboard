"""Lắp ráp quyển hồ sơ chất lượng hoàn chỉnh: bìa, trang xác nhận, mục lục có số trang, tờ phân cách,
các biểu mẫu (phiếu yêu cầu, biên bản, checklist), đánh số trang và bookmark.

Dùng: python tools/assemble_ho_so.py <thư_mục_ho_so>
Cần: soffice, pdftotext, pypdf, reportlab, python-docx.
"""
import datetime as dt
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).parent.parent
HS = Path(sys.argv[1])
DB = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))
PRJ, RECS = DB["project"], DB["records"]
TV = "Công ty TNHH Apave Châu Á – Thái Bình Dương"
NAVY = RGBColor(0x12, 0x3E, 0x70)
FD = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("Serif", FD + "LiberationSerif-Regular.ttf"))
pdfmetrics.registerFont(TTFont("SerifB", FD + "LiberationSerif-Bold.ttf"))
pdfmetrics.registerFont(TTFont("SerifI", FD + "LiberationSerif-Italic.ttf"))
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]
SUFFIXES = ("", "_ep_coc", "_cat_dau_coc", "_van_khuon", "_bt_lot", "_thep", "_khe_lun")
SUF_LABEL = {"_ep_coc": " (ép cọc)", "_cat_dau_coc": " (cắt đầu cọc)", "_van_khuon": " (ván khuôn)", "_bt_lot": " (đổ bê tông lót)",
             "_thep": " (cốt thép)", "_khe_lun": " (khe lún)"}
SHORT = {"15": "Đắp đê quai, tạo sàn thi công (độ chặt K95)", "16": "Đào bùn, nạo vét lạch phía núi",
         "04+05": "Đào móng, vận chuyển đất đào", "10+08": "Ván khuôn, đổ bê tông lót M100",
         "13+11+12+14": "Thép, ván khuôn (cấu kiện, khóa đỉnh), khe lún trước đổ bê tông", "09": "Đổ bê tông M300 bền sunfat, chống thấm",
         "06": "Đắp hoàn trả hố móng, lu nền (độ chặt K95)", "07": "Đắp đá xô chân tường chắn", "17": "Biện pháp thi công trọn gói",
         "02+03": "Nghiệm thu cọc: ép cọc, cắt đầu cọc"}
dmy = lambda s: dt.date.fromisoformat(s).strftime("%d/%m/%Y")
num = lambda r: r["so"].split()[-1]


# ------------------------------------------------------------------ docx helpers
def _font(run, size, bold=False, italic=False, color=None):
    run.font.name = "Times New Roman"
    run._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold, run.italic = bold, italic
    if color is not None:
        run.font.color.rgb = color


def _para(doc, text="", size=13, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.CENTER, color=None, before=0, after=0, keep=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(before), Pt(after)
    if text:
        _font(p.add_run(text), size, bold, italic, color)
    return p


def _border_bottom(p, color="123E70", sz=12):
    ppr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    e = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", str(sz)), ("space", "1"), ("color", color)):
        e.set(qn(f"w:{k}"), v)
    b.append(e)
    ppr.append(b)


def _page_border(section):
    pg = OxmlElement("w:pgBorders")
    pg.set(qn("w:offsetFrom"), "page")
    for side, val, sz in (("top", "double", "6"), ("left", "double", "6"), ("bottom", "double", "6"), ("right", "double", "6")):
        e = OxmlElement(f"w:{side}")
        for k, v in (("val", val), ("sz", sz), ("space", "24"), ("color", "123E70")):
            e.set(qn(f"w:{k}"), v)
        pg.append(e)
    pgmar = section._sectPr.find(qn("w:pgMar"))
    pgmar.addnext(pg)


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), fill)
    tcpr.append(sh)


def _cell_borders(table, color="9AA5B1", sz=4):
    tbl = table._tbl
    pr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        for k, v in (("val", "single"), ("sz", str(sz)), ("space", "0"), ("color", color)):
            e.set(qn(f"w:{k}"), v)
        b.append(e)
    pr.append(b)


def _a4(doc, top=2.0, bottom=2.0, left=2.5, right=1.5):
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Cm(top), Cm(bottom), Cm(left), Cm(right)
    return s


def _set_cell(cell, text, size=12, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
    _font(p.add_run(text), size, bold, italic, color)


# ------------------------------------------------------------------ bìa + trang xác nhận
def make_bia(total_pages, out):
    doc = Document()
    s = _a4(doc, 2.2, 2.0, 2.5, 2.0)
    _page_border(s)
    doc.styles["Normal"].font.name = "Times New Roman"
    p = _para(doc, after=6)
    p.add_run().add_picture(str(ROOT / "data/qaqc/templates/logo_giang_long_lg.png"), width=Cm(3.8))
    _para(doc, PRJ["nha_thau"].upper(), 17, True, color=NAVY, after=4)
    _border_bottom(_para(doc, after=14))
    _para(doc, "HỒ SƠ QUẢN LÝ CHẤT LƯỢNG", 30, True, color=NAVY, before=18, after=6)
    _para(doc, "NGHIỆM THU VẬT LIỆU, CÔNG VIỆC VÀ GIAI ĐOẠN", 14, True, after=0)
    _para(doc, "THI CÔNG XÂY DỰNG", 14, True, after=10)
    _border_bottom(_para(doc, after=18))
    _para(doc, "GÓI THẦU", 12, False, True, after=2)
    _para(doc, PRJ["goi_thau"].upper(), 16, True, color=NAVY, after=12)
    _para(doc, "DỰ ÁN", 12, False, True, after=2)
    _para(doc, PRJ["ten_cong_trinh"], 14, True, after=22)
    info = [("Hợp đồng số", PRJ["so_hop_dong"]), ("Chủ đầu tư", PRJ["chu_dau_tu"]), ("TVQLDA & GSTC", TV),
            ("Nhà thầu thi công", PRJ["nha_thau"]), ("Địa điểm xây dựng", PRJ["dia_diem"]),
            ("Phạm vi", "Tuyến 3, từ cọc 88 đến cọc 126"), ("Thời gian thi công", "12/08/2026 – 27/10/2026 (theo hợp đồng)"),
            ("Tập hồ sơ", "Quyển 1/1")]
    t = doc.add_table(rows=len(info), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, (a, b) in enumerate(info):
        t.rows[i].cells[0].width, t.rows[i].cells[1].width = Cm(4.2), Cm(11.8)
        _set_cell(t.rows[i].cells[0], a, 12, True, color=NAVY)
        _set_cell(t.rows[i].cells[1], b, 12)
    t.columns[0].width, t.columns[1].width = Cm(4.2), Cm(11.8)
    _cell_borders(t, "C5CCD6", 4)
    _para(doc, "Quảng Ninh, năm 2026", 13, False, True, before=26)
    # trang xác nhận
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    _para(doc, "XÁC NHẬN HỒ SƠ", 20, True, color=NAVY, before=10, after=4)
    _border_bottom(_para(doc, after=14))
    _para(doc, f"Quyển hồ sơ gồm {total_pages} trang (kể cả bìa, trang xác nhận và mục lục), được lập đầy đủ theo danh mục kèm theo.", 13,
          align=WD_ALIGN_PARAGRAPH.LEFT, after=6)
    _para(doc, "Kết quả kiểm tra, kết luận nghiệm thu, ngày tháng và chữ ký thực tế do các bên trực tiếp tham gia nghiệm thu ghi nhận trong từng biểu mẫu.",
          13, align=WD_ALIGN_PARAGRAPH.LEFT, after=24)
    sg = doc.add_table(rows=2, cols=2)
    sg.autofit = False
    heads = [("NHÀ THẦU THI CÔNG", PRJ["nha_thau"], "Chỉ huy trưởng", PRJ["chi_huy_truong"]),
             ("TVQLDA & GSTC", TV, "Cán bộ kỹ thuật", PRJ["can_bo_tvgs"])]
    for c, (h, org, role, name) in enumerate(heads):
        cell = sg.rows[0].cells[c]
        cell.width = Cm(8.0)
        cell.text = ""
        for text, sz, b in ((h, 13, True), (org, 12, False), (f"{role}", 12, False)):
            p = cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _font(p.add_run(text), sz, b)
        cell2 = sg.rows[1].cells[c]
        cell2.width = Cm(8.0)
        cell2.text = ""
        p = cell2.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        for _ in range(7):
            r.add_break()
        r.add_text(name)
        _font(r, 13, True)
    _para(doc, "Ngày ........ tháng ........ năm ..........", 12, False, True, before=14)
    doc.save(out)


# ------------------------------------------------------------------ mục lục
def make_muc_luc(rows, out):
    doc = Document()
    _a4(doc, 1.8, 1.8, 2.5, 1.5)
    doc.styles["Normal"].font.name = "Times New Roman"
    _para(doc, "DANH MỤC HỒ SƠ", 20, True, color=NAVY, after=2)
    _para(doc, f"{PRJ['goi_thau']} — Hợp đồng số {PRJ['so_hop_dong']}", 11, False, True, after=2)
    _border_bottom(_para(doc, after=8))
    widths = [1.2, 1.9, 5.2, 2.5, 1.7, 1.7, 2.8]
    t = doc.add_table(rows=1, cols=7)
    t.autofit = False
    _cell_borders(t)
    for i, h in enumerate(["STT", "Số hiệu", "Nội dung", "Ngày nghiệm thu", "Phiếu yêu cầu", "Biên bản", "Kèm theo"]):
        c = t.rows[0].cells[i]
        c.width = Cm(widths[i])
        _set_cell(c, h, 10.5, True, align=WD_ALIGN_PARAGRAPH.CENTER, color=RGBColor(255, 255, 255))
        _shade(c, "123E70")
    for i, wcm in enumerate(widths):
        t.columns[i].width = Cm(wcm)
    trpr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trpr.append(th)
    for kind, data in rows:
        r = t.add_row()
        for i, c in enumerate(r.cells):
            c.width = Cm(widths[i])
        trpr = r._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        if kind in ("sec", "sub"):
            m = r.cells[0].merge(r.cells[6])
            _set_cell(m, data, 11.5 if kind == "sec" else 11, True, align=WD_ALIGN_PARAGRAPH.LEFT, color=NAVY)
            _shade(m, "D9E2F3" if kind == "sec" else "EEF2F8")
        else:
            for i, v in enumerate(data):
                _set_cell(r.cells[i], str(v), 11, False, align=WD_ALIGN_PARAGRAPH.LEFT if i == 2 else WD_ALIGN_PARAGRAPH.CENTER)
    doc.save(out)


def write_danh_muc_xlsx(rows, out):
    """Danh mục hồ sơ dạng Excel (có số trang), kèm sheet chi tiết cũ nếu có."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    old = None
    if out.exists():
        try:
            old = pd.read_excel(out, sheet_name=0)
            if "Số" not in old.columns:
                old = None
        except Exception:
            old = None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Danh_muc_ho_so"
    thin = Side(style="thin", color="9AA5B1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    navy = PatternFill("solid", fgColor="123E70")
    ws.merge_cells("A1:G1")
    ws["A1"] = "DANH MỤC HỒ SƠ QUẢN LÝ CHẤT LƯỢNG"
    ws["A1"].font = Font(name="Times New Roman", size=16, bold=True, color="123E70")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:G2")
    ws["A2"] = f"{PRJ['goi_thau']} - Hợp đồng số {PRJ['so_hop_dong']} - {PRJ['nha_thau']}"
    ws["A2"].font = Font(name="Times New Roman", size=11, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[2].height = 32
    hdr = ["STT", "Số hiệu", "Nội dung", "Ngày nghiệm thu", "Phiếu yêu cầu (trang)", "Biên bản (trang)", "Kèm theo"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(4, i, h)
        c.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        c.fill = navy
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[4].height = 32
    r = 5
    for kind, data in rows:
        if kind in ("sec", "sub"):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            c = ws.cell(r, 1, data)
            c.font = Font(name="Times New Roman", size=11, bold=True, color="123E70")
            c.fill = PatternFill("solid", fgColor="D9E2F3" if kind == "sec" else "EEF2F8")
            for col in range(1, 8):
                ws.cell(r, col).border = border
        else:
            for i, v in enumerate(data, 1):
                c = ws.cell(r, i, v)
                c.font = Font(name="Times New Roman", size=11)
                c.border = border
                c.alignment = Alignment(horizontal="left" if i == 3 else "center", vertical="center", wrap_text=True)
        r += 1
    for col, wd in zip("ABCDEFG", (6, 11, 58, 15, 14, 12, 18)):
        ws.column_dimensions[col].width = wd
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "4:4"
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = 0.6, 0.6, 0.7, 0.7
    if old is not None:
        ws2 = wb.create_sheet("Chi_tiet")
        for j, col in enumerate(old.columns, 1):
            c = ws2.cell(1, j, col)
            c.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
            c.fill = navy
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, row in enumerate(old.itertuples(index=False), 2):
            for j, v in enumerate(row, 1):
                ws2.cell(i, j, None if pd.isna(v) else v).alignment = Alignment(wrap_text=True, vertical="top")
        for j, col in enumerate(old.columns, 1):
            ws2.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 14 if col != "Hạng mục" else 55
        ws2.freeze_panes = "A2"
    wb.save(out)


# ------------------------------------------------------------------ PDF helpers
def to_pdf(files, outdir):
    outdir.mkdir(parents=True, exist_ok=True)
    for i in range(0, len(files), 15):
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir)] + [str(f) for f in files[i:i + 15]],
                       check=True, capture_output=True, timeout=900)


def page_texts(pdf):
    out = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True).stdout
    return out.split("\f")


def find_pages(pdf, pattern_fn, records):
    """Trang bắt đầu (0-based) của từng bản ghi trong file PDF gộp nhiều biểu mẫu."""
    texts = page_texts(pdf)
    res = {}
    for r in records:
        pat = re.compile(pattern_fn(r))
        for i, tx in enumerate(texts):
            if pat.search(tx.replace("\xa0", " ")) and r["so"] not in res:
                res[r["so"]] = i
                break
    return res


def divider(title, subtitle, items, letter):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    c.setStrokeColorRGB(0.07, 0.24, 0.44)
    c.setLineWidth(2)
    c.rect(40, 40, w - 80, h - 80)
    c.setLineWidth(0.6)
    c.rect(46, 46, w - 92, h - 92)
    c.setFillColorRGB(0.07, 0.24, 0.44)
    c.setFont("SerifB", 60)
    c.drawCentredString(w / 2, h * 0.62, letter)
    c.setFont("SerifB", 22)
    y = h * 0.54
    for line in title:
        c.drawCentredString(w / 2, y, line)
        y -= 28
    c.setFont("SerifI", 13)
    c.setFillColorRGB(0.25, 0.25, 0.25)
    c.drawCentredString(w / 2, y - 6, subtitle)
    c.setFont("Serif", 12)
    y -= 50
    for it in items:
        c.drawCentredString(w / 2, y, it)
        y -= 18
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp(page, text_l, text_r, w, h):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    c.setFont("Serif", 8.5)
    c.setFillColorRGB(0.35, 0.35, 0.35)
    c.drawString(60, 20, text_l)
    c.drawRightString(w - 40, 20, text_r)
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.setLineWidth(0.4)
    c.line(60, 32, w - 40, 32)
    c.save()
    buf.seek(0)
    page.merge_page(PdfReader(buf).pages[0])


# ------------------------------------------------------------------ main
def main():
    tmp = Path(tempfile.mkdtemp())
    pdfd = tmp / "pdf"
    docs = {"vl": HS / "02_Nghiem_thu_vat_lieu_dau_vao.docx", "lm": HS / "06_Bien_ban_lay_mau.docx",
            "yc_vl": HS / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_vat_lieu.docx",
            "gd": HS / "03_Nghiem_thu_giai_doan_coc.docx", "yc_gd": HS / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_giai_doan_coc.docx",
            "th": HS / "03_Bang_tong_hop_thi_cong_coc.xlsx"}
    for i, p in enumerate(PD, 1):
        docs[f"cv{i}"] = next(HS.glob(f"03_Nghiem_thu_cong_viec_PD{i}_*.docx"))
        docs[f"yc{i}"] = next((HS / "05_Phieu_yeu_cau_nghiem_thu").glob(f"YC_nghiem_thu_cong_viec_PD{i}_*.docx"))
    docs["cv0"] = HS / "03_Nghiem_thu_cong_viec_toan_tuyen.docx"
    docs["yc0"] = HS / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_cong_viec_toan_tuyen.docx"
    import openpyxl
    wb_tn = openpyxl.load_workbook(HS / "07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx")
    for n_ in wb_tn.sheetnames[2:]:
        del wb_tn[n_]
    tn_print = tmp / "07_Bang_theo_doi_ket_qua_thi_nghiem.xlsx"
    wb_tn.save(tn_print)
    docs["tn"] = tn_print
    cl_files = sorted((HS / "04_Checklist").rglob("*.xlsx"))
    to_pdf(list(docs.values()) + cl_files, pdfd)
    P = {k: pdfd / (v.stem + ".pdf") for k, v in docs.items()}
    CL = {f.stem: pdfd / (f.stem + ".pdf") for f in cl_files}

    vl = sorted([r for r in RECS if r["type"] == "VL"], key=lambda r: (r["ngay"], r["so"]))
    cv_all = sorted([r for r in RECS if r["type"] == "CV"], key=lambda r: (r["ngay"], r["so"]))
    gd = next(r for r in RECS if r["type"] == "GD")
    groups = [(f"Phân đoạn {i} - {p.replace('Cọc ', 'cọc ').replace('-', ' đến ')}", p, f"{i}") for i, p in enumerate(PD, 1)]
    groups.append(("Toàn tuyến", None, "0"))
    cv_of = {key: [r for r in cv_all if (r["vi_tri"] == p if p else r["vi_tri"] not in PD)] for _, p, key in groups}

    # --- thứ tự các phần (parts), số trang tính sau
    parts = []  # (level_title, section_key, pdf_path or page object, label)
    secs = [("A", ["NGHIỆM THU", "VẬT LIỆU ĐẦU VÀO"], "Phiếu yêu cầu, biên bản lấy mẫu và biên bản nghiệm thu vật liệu"),
            ("B", ["NGHIỆM THU", "CÔNG VIỆC XÂY DỰNG"], "Theo từng phân đoạn: phiếu yêu cầu, biên bản và checklist kiểm tra"),
            ("C", ["NGHIỆM THU GIAI ĐOẠN", "THI CÔNG CỌC"], "Phiếu yêu cầu, biên bản nghiệm thu giai đoạn và bảng tổng hợp thi công cọc"),
            ("D", ["TÀI LIỆU KÈM THEO"], "Chứng chỉ vật liệu, kết quả thí nghiệm, bản vẽ hoàn công, nhật ký thi công")]
    plan = []  # list of (kind, payload)
    plan.append(("div", "A"))
    plan += [("pdf", ("Phiếu yêu cầu nghiệm thu vật liệu", P["yc_vl"], "yc_vl")), ("pdf", ("Biên bản lấy mẫu thí nghiệm", P["lm"], "lm")),
             ("pdf", ("Biên bản nghiệm thu vật liệu đầu vào", P["vl"], "vl"))]
    plan.append(("div", "B"))
    for title, p, key in groups:
        plan.append(("pdf", (f"{title} - Phiếu yêu cầu nghiệm thu", P[f"yc{key}"], f"yc{key}")))
        plan.append(("pdf", (f"{title} - Biên bản nghiệm thu công việc", P[f"cv{key}"], f"cv{key}")))
        for r in cv_of[key]:
            for suf in SUFFIXES:
                stem = f"Checklist_{r['so'].replace(' ', '_')}{suf}"
                if stem in CL:
                    plan.append(("pdf", (f"Checklist {r['so']} - {SHORT[r['muc_hd']]}" + SUF_LABEL.get(suf, ""), CL[stem], stem)))
    plan.append(("div", "C"))
    plan += [("pdf", ("Phiếu yêu cầu nghiệm thu giai đoạn cọc", P["yc_gd"], "yc_gd")), ("pdf", ("Biên bản nghiệm thu giai đoạn cọc", P["gd"], "gd")),
             ("pdf", ("Bảng tổng hợp công tác thi công cọc", P["th"], "th"))]
    plan.append(("div", "D"))
    plan.append(("pdf", ("Danh sách kết quả thí nghiệm và hồ sơ kèm theo cần có", P["tn"], "tn")))

    # --- số trang
    def build_rows(start_of, n_front):
        """start_of: dict key->trang (1-based) đầu của file; trả các dòng mục lục."""
        pg = lambda k: f"{k}"
        rows = []
        rows.append(("sec", "PHẦN A - NGHIỆM THU VẬT LIỆU ĐẦU VÀO"))
        marks = {"yc_vl": find_pages(P["yc_vl"], lambda r: rf"Phiếu yêu cầu số {num(r)}/YC/VL-TB", vl),
                 "lm": find_pages(P["lm"], lambda r: rf"Biên bản số {r['so_lm'].split()[-1]}/LM", [x for x in RECS if x.get("lay_mau")]),
                 "vl": find_pages(P["vl"], lambda r: rf"Biên bản số {num(r)}/CV/VL-TB", vl)}
        for i, r in enumerate(vl, 1):
            rows.append(("item", (i, r["so"], r["hang_muc"] if "Bê tông" not in r["hang_muc"] else f"Bê tông thương phẩm M300 bền sunfat - {r['vi_tri'].lower()}",
                                  dmy(r["ngay"]), start_of["yc_vl"] + marks["yc_vl"][r["so"]], start_of["vl"] + marks["vl"][r["so"]],
                                  (f"LM tr. {start_of['lm'] + marks['lm'][r['so']]}" if r.get("lay_mau") else "-"))))
        rows.append(("sec", "PHẦN B - NGHIỆM THU CÔNG VIỆC XÂY DỰNG"))
        stt = 0
        for title, p, key in groups:
            rows.append(("sub", title))
            ycm = find_pages(P[f"yc{key}"], lambda r: rf"Phiếu yêu cầu số {num(r)}/YC/CV", cv_of[key])
            bbm = find_pages(P[f"cv{key}"], lambda r: rf"Biên bản số {num(r)}/CV(?!/)", cv_of[key])
            for r in cv_of[key]:
                stt += 1
                cks = [start_of.get(f"Checklist_{r['so'].replace(' ', '_')}{s}") for s in SUFFIXES]
                cks = [c for c in cks if c]
                ck = f"tr. {cks[0]}" if cks else "-"
                if len(cks) == 2:
                    ck = f"tr. {cks[0]}, {cks[1]}"
                if r.get("lay_mau"):  # đổ bê tông: kèm biên bản lấy mẫu (thí nghiệm R7, R28)
                    lm_p = start_of["lm"] + marks["lm"][r["so"]]
                    ck = f"LM tr. {lm_p}" + (f"; CL {ck[3:]}" if cks else "")
                rows.append(("item", (stt, r["so"], SHORT[r["muc_hd"]] + (f" - {r['vi_tri']}" if not p else ""), dmy(r["ngay"]),
                                      start_of[f"yc{key}"] + ycm[r["so"]], start_of[f"cv{key}"] + bbm[r["so"]], ck)))
        rows.append(("sec", "PHẦN C - NGHIỆM THU GIAI ĐOẠN THI CÔNG CỌC"))
        ycg = find_pages(P["yc_gd"], lambda r: rf"Phiếu yêu cầu số {num(r)}/YC/GĐ", [gd])
        bbg = find_pages(P["gd"], lambda r: rf"Biên bản số {num(r)}/CV/GĐ", [gd])
        rows.append(("item", (1, gd["so"], "Nghiệm thu hoàn thành giai đoạn thi công cọc toàn tuyến", dmy(gd["ngay"]), start_of["yc_gd"] + ycg[gd["so"]],
                              start_of["gd"] + bbg[gd["so"]], f"Bảng tổng hợp tr. {start_of['th']}")))
        rows.append(("sec", "PHẦN D - TÀI LIỆU KÈM THEO (BỔ SUNG)"))
        tn = pd.read_excel(HS / "07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx", sheet_name="Danh_sach", header=3)
        rows.append(("item", (1, "-", "Bảng theo dõi kết quả thí nghiệm và hồ sơ cần có", "-", "-", "-", f"tr. {start_of['tn']}")))
        full = {"A": "Cọc BT300x300 (vật liệu đầu vào)", "B": "Cốt thép", "C": "Bê tông M300 theo từng ngày đổ (R7, R28)",
                "D": "Đắp đất - độ chặt K95", "E": "Đá xô", "F": "Nén tĩnh cọc (nếu thiết kế yêu cầu)"}
        for j, (nhom, cnt) in enumerate(tn["Nhóm"].value_counts(sort=False).items(), 2):
            rows.append(("item", (j, "-", f"{nhom[0]}. {full[nhom[0]]} ({cnt} thí nghiệm)", "-", "-", "-", "Bổ sung")))
        hs = pd.read_excel(HS / "07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx", sheet_name="Ho_so_kem_theo", header=3)
        rows.append(("item", (len(tn["Nhóm"].unique()) + 2, "-", f"Hồ sơ kèm theo: chứng chỉ, cấp phối, phiếu giao hàng, nhật ký ({len(hs)} mục)", "-", "-", "-", "Bổ sung")[:7]))
        return rows

    # lượt 1: số trang giả để biết số trang mục lục
    plan_pages = []
    for kind, pl in plan:
        plan_pages.append(1 if kind == "div" else len(PdfReader(str(pl[1])).pages))
    dummy = {pl[2]: 999 for kind, pl in plan if kind == "pdf"}
    make_muc_luc(build_rows(dummy, 0), tmp / "ml.docx")
    to_pdf([tmp / "ml.docx"], tmp / "mlpdf")
    n_ml = len(PdfReader(str(tmp / "mlpdf/ml.pdf")).pages)
    front = 2 + n_ml
    start_of, cur = {}, front + 1
    for (kind, pl), n in zip(plan, plan_pages):
        if kind == "pdf":
            start_of[pl[2]] = cur
        cur += n
    total = cur - 1
    rows = build_rows(start_of, front)
    make_muc_luc(rows, HS / "00_Muc_luc_ho_so.docx")
    write_danh_muc_xlsx(rows, HS / "00_Danh_muc_ho_so.xlsx")
    make_bia(total, HS / "00_Bia_ho_so.docx")
    to_pdf([HS / "00_Muc_luc_ho_so.docx", HS / "00_Bia_ho_so.docx"], tmp / "front")
    bia = PdfReader(str(tmp / "front/00_Bia_ho_so.pdf"))
    ml = PdfReader(str(tmp / "front/00_Muc_luc_ho_so.pdf"))
    assert len(ml.pages) == n_ml and len(bia.pages) == 2, (len(ml.pages), len(bia.pages))

    # --- ghép
    w = PdfWriter()
    for p in bia.pages:
        w.add_page(p)
    for p in ml.pages:
        w.add_page(p)
    b_bia = w.add_outline_item("Bìa và trang xác nhận", 0)
    w.add_outline_item("Danh mục hồ sơ", 2)
    cur_parent = None
    for (kind, pl), n in zip(plan, plan_pages):
        if kind == "div":
            letter = pl
            _, title, sub = next(s for s in secs if s[0] == letter)
            items = {"A": ["Phiếu yêu cầu nghiệm thu vật liệu", "Biên bản lấy mẫu thí nghiệm", "Biên bản nghiệm thu vật liệu đầu vào"],
                     "B": ["Phân đoạn 1 - 4 và toàn tuyến", "Phiếu yêu cầu, biên bản nghiệm thu công việc, checklist"],
                     "C": ["Phiếu yêu cầu nghiệm thu giai đoạn", "Biên bản nghiệm thu giai đoạn", "Bảng tổng hợp thi công cọc"],
                     "D": ["Bảng theo dõi kết quả thí nghiệm cần có (trang sau)", "Đặt tại đây theo thứ tự bảng: chứng chỉ vật liệu, phiếu kết quả", "thí nghiệm, bản vẽ hoàn công, nhật ký thi công"]}[letter]
            w.add_page(divider(title, sub, items, f"PHẦN {letter}"))
            cur_parent = w.add_outline_item(f"Phần {letter} - {' '.join(title).title() if False else ' '.join(title)}", len(w.pages) - 1)
        else:
            label, path, key = pl
            first = len(w.pages)
            for p in PdfReader(str(path)).pages:
                w.add_page(p)
            w.add_outline_item(label, first, parent=cur_parent)
    # --- đánh số trang + chân trang
    N = len(w.pages)
    assert N == total, (N, total)
    for i, page in enumerate(w.pages):
        if i < 1:
            continue
        pw, ph = float(page.mediabox.width), float(page.mediabox.height)
        stamp(page, f"Hồ sơ quản lý chất lượng - {PRJ['nha_thau']} - Hợp đồng {PRJ['so_hop_dong']}", f"Trang {i + 1}/{N}", pw, ph)
    out = HS / "00_HO_SO_HOAN_CHINH.pdf"
    with open(out, "wb") as f:
        w.write(f)
    print("OK", N, "trang;", n_ml, "trang mục lục;", out.stat().st_size // 1024, "KB")
    shutil.rmtree(tmp, ignore_errors=True)


main()
