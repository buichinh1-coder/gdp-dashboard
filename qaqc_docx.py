"""Xuất biên bản nghiệm thu (.docx) bằng cách nhân bản form mẫu của TVGS.

Chỉ điền dữ liệu người dùng đã nhập; chỗ chưa có dữ liệu giữ nguyên dấu chấm
để điền tay. Không tự điền kết luận "Đạt".
"""
import copy
import io
from datetime import date
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.table import Table

TEMPLATE = Path(__file__).parent / "data" / "qaqc" / "templates" / "mau_bien_ban.docx"

# Khoảng phần tử body trong file mẫu (chỉ số bắt đầu, kết thúc) của từng form.
FORMS = {
    "VL": (108, 156),  # Biên bản /CV/VL-TB
    "CV": (184, 234),  # Biên bản /CV
}
SUFFIX = {"VL": "CV/VL-TB", "CV": "CV"}


def _set_text(p_el, text):
    """Ghi lại toàn bộ đoạn, giữ định dạng của run đầu tiên."""
    ts = p_el.findall('.//' + qn('w:t'))
    if not ts:
        return
    ts[0].text = text
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for t in ts[1:]:
        t.text = ''


def _text(p_el):
    return ''.join(t.text or '' for t in p_el.findall('.//' + qn('w:t')))


def _num(rec):
    return rec["so"].split(" ", 1)[-1]


def _d(iso):
    d = date.fromisoformat(iso)
    return d.day, d.month, d.year


def build_docx(records, project):
    """records: list biên bản loại VL/CV. Trả về bytes của file .docx."""
    doc = docx.Document(str(TEMPLATE))
    body = doc.element.body
    src = list(body)
    sect = src[-1]
    for el in src[:-1]:
        body.remove(el)

    for n, rec in enumerate(records):
        lo, hi = FORMS[rec["type"]]
        els = [copy.deepcopy(src[i]) for i in range(lo, hi + 1)]
        for el in els:
            sect.addprevious(el)
        P = lambda i: els[i - lo]
        dd, mm, yy = _d(rec["ngay"])
        ngay = f"ngày {dd:02d} tháng {mm:02d} năm {yy}"
        if n > 0 or True:  # mỗi biên bản bắt đầu trang mới
            first = els[0]
            ppr = first.find(qn('w:pPr'))
            if ppr is None:
                ppr = first.makeelement(qn('w:pPr'), {})
                first.insert(0, ppr)
            ppr.append(ppr.makeelement(qn('w:pageBreakBefore'), {}))

        # Các đoạn dùng chung: ngày, số, dự án, địa điểm, gói thầu.
        for i in range(lo, hi + 1):
            s = _text(els[i - lo]).strip()
            if s.startswith(("Quảng Ninh, ngày", "Biên bản số")):
                if s.startswith("Biên bản số"):
                    _set_text(els[i - lo], f"Biên bản số {_num(rec)}/{SUFFIX[rec['type']]}")
                else:
                    _set_text(els[i - lo], f"Quảng Ninh, {ngay}")
            elif s.startswith("Dự án:") and project.get("ten_cong_trinh"):
                _set_text(els[i - lo], f"Dự án: {project['ten_cong_trinh']}")
            elif s.startswith("Địa điểm"):  # không giữ địa điểm của dự án mẫu
                _set_text(els[i - lo], f"Địa điểm       :  {project.get('dia_diem') or '……………………………………………'}")
            elif s.startswith("Gói thầu:") and project.get("goi_thau"):
                _set_text(els[i - lo], f"Gói thầu: {project['goi_thau']}")
            elif s.startswith("1. Đối tượng nghiệm thu") and rec.get("hang_muc"):
                vitri = f" - {rec['vi_tri']}" if rec.get("vi_tri") else ""
                kl = f" - khối lượng: {rec['khoi_luong']:g} {rec.get('dvt', '')}".rstrip() if rec["type"] == "CV" and rec.get("khoi_luong") else ""
                _set_text(els[i - lo], f"1. Đối tượng nghiệm thu : {rec['hang_muc']}{vitri}{kl}")
            elif s.startswith("Đại diện nhà thầu thi công xây dựng công trình") and project.get("nha_thau"):
                _set_text(els[i - lo], f"Đại diện nhà thầu thi công xây dựng công trình: {project['nha_thau']}")
            elif s.startswith(("Bắt đầu", "Kết thúc")):
                label = "Bắt đầu  :" if s.startswith("Bắt đầu") else "Kết thúc :"
                _set_text(els[i - lo], f"{label}   {ngay}")
            elif s.startswith("5. Kết luận") or s.startswith("b. Đánh giá"):
                pass

        # Kết luận: chỉ điền khi người kiểm tra đã nhập.
        kl = rec.get("ket_luan", "")
        if kl and kl != "Chưa có kết luận":
            for i in range(lo, hi + 1):
                if _text(els[i - lo]).strip().startswith("5. Kết luận"):
                    _set_text(els[i + 1 - lo], kl)

        # Bảng khối lượng của form VL.
        if rec["type"] == "VL":
            for el in els:
                if el.tag == qn('w:tbl'):
                    t = Table(el, doc)
                    if t.rows[0].cells[0].text.strip() == "STT" and len(t.rows) > 1:
                        vals = ["1", rec.get("hang_muc", ""), "", rec.get("dvt", ""),
                                f"{rec.get('khoi_luong', 0):g}" if rec.get("khoi_luong") else ""]
                        for c, v in zip(t.rows[1].cells, vals):
                            if v:
                                c.paragraphs[0].add_run(v)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
