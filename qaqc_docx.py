"""Xuất biểu mẫu nghiệm thu (.docx) bằng cách nhân bản form mẫu của TVGS.

Chỉ điền dữ liệu người dùng đã nhập; chỗ chưa có dữ liệu giữ nguyên dấu chấm
để điền tay. Không tự điền kết luận "Đạt".

Loại biểu mẫu:
  VL   Biên bản nghiệm thu vật liệu (/CV/VL-TB)     CV   Biên bản nghiệm thu công việc (/CV)
  LM   Biên bản lấy mẫu thí nghiệm (/LM)            YCVL Phiếu yêu cầu nghiệm thu vật liệu (/YC/VL-TB)
  YCCV Phiếu yêu cầu nghiệm thu công việc (/YC/CV)
"""
import copy
import io
import re
from datetime import date
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table

TEMPLATE = Path(__file__).parent / "data" / "qaqc" / "templates" / "mau_bien_ban.docx"

# Khoảng phần tử body trong file mẫu (chỉ số bắt đầu, kết thúc) của từng form.
FORMS = {
    "LM": (43, 81),
    "YCVL": (82, 107),
    "VL": (108, 156),
    "YCCV": (158, 183),
    "CV": (184, 234),
    "YCGD": (497, 520),  # phiếu yêu cầu nghiệm thu giai đoạn
    "GD": (521, 579),    # biên bản nghiệm thu giai đoạn
}
SUFFIX = {"VL": "CV/VL-TB", "CV": "CV", "LM": "LM", "YCVL": "YC/VL-TB", "YCCV": "YC/CV", "GD": "CV/GĐ", "YCGD": "YC/GĐ"}
IS_YC = ("YCVL", "YCCV", "YCGD")


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


def _kl(x):
    return f"{x:g}".replace(".", ",")


def _doi_tuong(rec):
    vitri = f" - {rec['vi_tri']}" if rec.get("vi_tri") else ""
    kl = ""
    if rec["type"] in ("VL", "YCVL") and rec.get("khoi_luong"):  # chỉ vật liệu đầu vào mới ghi khối lượng
        kl = f" - khối lượng: {_kl(rec['khoi_luong'])} {rec.get('dvt', '')}".rstrip()
    return f"{rec['hang_muc']}{vitri}{kl}"


FONT = "Times New Roman"
SIZE = Pt(13)


def _chuan_font(el):
    """Mọi chữ dùng Times New Roman, cỡ tối thiểu 13pt (tiêu đề lớn hơn thì giữ)."""
    for r in el.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = r.makeelement(qn('w:rPr'), {})
            r.insert(0, rpr)
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = rpr.makeelement(qn('w:rFonts'), {})
            rpr.insert(0, rf)
        for a in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(qn(f"w:{a}"), FONT)
        for tag in ("w:sz", "w:szCs"):
            sz = rpr.find(qn(tag))
            if sz is None:
                sz = rpr.makeelement(qn(tag), {})
                rpr.append(sz)
            cur = int(sz.get(qn("w:val")) or 0)
            sz.set(qn("w:val"), str(max(cur, 26)))  # 26 half-point = 13pt


def _name_in_cell(cell, name, blanks=3):
    """Thêm tên in dưới ô ký (chừa chỗ ký phía trên)."""
    p = cell.paragraphs[0]
    r = p.add_run()
    for _ in range(blanks):
        r.add_break()
    r.add_text(name)
    r.bold = True
    r.font.name = "Times New Roman"
    p.alignment = 1


def _gon_o(cell):
    """Giữ các đoạn có chữ, bỏ đoạn trống; căn giữa để hai ô ký đối xứng."""
    ps = cell.paragraphs
    keep = [p for p in ps if p.text.strip()] or ps[:1]
    for p in ps:
        if p not in keep and len(cell.paragraphs) > 1:
            p._p.getparent().remove(p._p)
    for p in cell.paragraphs:
        p.alignment = 1


def _ten_nguoi_ky(els, doc, t, project):
    from docx.enum.text import WD_TAB_ALIGNMENT
    tung, anh = project.get("chi_huy_truong"), project.get("can_bo_tvgs")
    for el in els:
        if el.tag != qn('w:tbl'):
            continue
        tb = Table(el, doc)
        head = tb.rows[0].cells[0].text.strip()
        if t == "LM" and head.startswith("Đại diện TVGS"):
            for c in tb.rows[0].cells:
                _gon_o(c)
            if anh:
                _name_in_cell(tb.rows[1].cells[0], anh, 3)
            if tung:
                _name_in_cell(tb.rows[1].cells[1], tung, 3)
        elif t == "CV" and head.startswith("Đại diện TVGS"):
            for c in tb.rows[0].cells:
                _gon_o(c)
            if anh:
                _name_in_cell(tb.rows[1].cells[0], anh, 2)
            if tung:
                _name_in_cell(tb.rows[1].cells[1], tung, 2)
        elif t in IS_YC and head.startswith(("Đơn vị TVQLDA", "Chủ đầu tư")):
            cells = tb.rows[1].cells
            pairs = [(cells[-2], anh, "Người nhận:"), (cells[-1], tung, None)] if t == "YCGD" else [(cells[0], anh, "Người nhận:"), (cells[1], tung, None)]
            if t == "YCGD":
                pairs = [(cells[1], anh, "Người nhận:")]
            for c, name, label in pairs:
                if not name:
                    continue
                first = c.paragraphs[0]
                for p in c.paragraphs[1:]:
                    p._p.getparent().remove(p._p)
                if label:
                    _set_text(first._p, label)
                _name_in_cell(c, name, 3)
        elif t == "GD" and head.startswith("Đại diện TVQLDA"):
            for c in tb.rows[0].cells:
                _gon_o(c)
            if anh:
                _name_in_cell(tb.rows[0].cells[0], anh, 4)
            if tung:
                _name_in_cell(tb.rows[0].cells[1], tung, 4)
        elif t == "VL" and head.startswith("Đại diện TVQLDA"):
            for c in tb.rows[0].cells:
                _gon_o(c)
            if anh and tung:
                cell = tb.rows[1].cells[0]
                ps = cell.paragraphs
                for p in ps[:4]:
                    p.alignment = 0
                    p.paragraph_format.left_indent = 0
                    p.paragraph_format.first_line_indent = 0
                    # tâm tab = tâm hai cột tiêu đề của bảng mẫu (cột 1: 7,94 cm; cột 2: 9,15 cm; trừ lề ô 0,19 cm)
                    w0, w1 = tb.columns[0].width.cm, tb.columns[1].width.cm
                    p.paragraph_format.tab_stops.add_tab_stop(Cm(w0 / 2 - 0.19), WD_TAB_ALIGNMENT.CENTER)
                    p.paragraph_format.tab_stops.add_tab_stop(Cm(w0 + w1 / 2 - 0.19), WD_TAB_ALIGNMENT.CENTER)
                    for r in list(p.runs):
                        r._r.getparent().remove(r._r)
                # ps[0..2]: chừa chỗ ký; ps[3]: họ tên người ký, bên dưới tên nhà thầu là chức danh (theo mẫu bạn gửi)
                r3 = ps[3].add_run(f"\t{anh}\t{tung}")
                r3.bold, r3.font.name = True, "Times New Roman"
                r3.add_break()
                r4 = ps[3].add_run("\t\tChỉ huy trưởng công trường")
                r4.italic, r4.font.name = True, "Times New Roman"


def build_docx(records, project):
    """records: list biên bản (type thuộc FORMS). Trả về bytes của file .docx."""
    doc = docx.Document(str(TEMPLATE))
    body = doc.element.body
    src = list(body)
    sect = src[-1]
    for el in src[:-1]:
        body.remove(el)
    sec = doc.sections[-1]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4 dọc
    sec.top_margin, sec.bottom_margin = Cm(1.5), Cm(1.5)
    sec.left_margin, sec.right_margin = Cm(2.5), Cm(1.5)
    sec.header_distance = sec.footer_distance = Cm(0.7)

    for rec in records:
        t = rec["type"]
        lo, hi = FORMS[t]
        els = [copy.deepcopy(src[i]) for i in range(lo, hi + 1)]
        if t == "LM":  # gọn trong 1 trang A4: bỏ dòng trống, giãn dòng vừa phải, không tách hàng ký
            els = [e for e in els if e.tag != qn('w:p') or _text(e).strip() or e.findall('.//' + qn('w:drawing'))]
            for e in els:  # giãn dòng đơn (không đổi cỡ chữ) để dành chỗ cho khoảng ký
                for pp in ([e] if e.tag == qn('w:p') else []):
                    ppr = pp.find(qn('w:pPr'))
                    if ppr is None:
                        ppr = pp.makeelement(qn('w:pPr'), {})
                        pp.insert(0, ppr)
                    sp = ppr.find(qn('w:spacing'))
                    if sp is None:
                        sp = ppr.makeelement(qn('w:spacing'), {})
                        ppr.append(sp)
                    sp.set(qn('w:line'), '240')
                    sp.set(qn('w:lineRule'), 'auto')
            for e in els:
                if e.tag == qn('w:tbl'):
                    for tr in e.iter(qn('w:tr')):
                        trpr = tr.find(qn('w:trPr'))
                        if trpr is None:
                            trpr = tr.makeelement(qn('w:trPr'), {})
                            tr.insert(0, trpr)
                        trpr.append(trpr.makeelement(qn('w:cantSplit'), {}))
        for el in els:
            for br in list(el.iter(qn('w:br'))):
                if br.get(qn('w:type')) == 'page':
                    br.getparent().remove(br)
            sect.addprevious(el)
        dd, mm, yy = _d(rec["ngay"])
        ngay = f"ngày {dd:02d} tháng {mm:02d} năm {yy}"
        first = els[0]  # mỗi biểu mẫu bắt đầu trang mới
        ppr = first.find(qn('w:pPr'))
        if ppr is None:
            ppr = first.makeelement(qn('w:pPr'), {})
            first.insert(0, ppr)
        ppr.append(ppr.makeelement(qn('w:pageBreakBefore'), {}))

        prefix = "Phiếu yêu cầu số" if t in IS_YC else "Biên bản số"
        n = len(els)
        people = {"tv": [(project.get("giam_sat_truong"), "Giám sát trưởng"), (project.get("can_bo_tvgs"), "Cán bộ kỹ thuật")],
                  "nt": [(project.get("chi_huy_truong"), "Chỉ huy trưởng")], "lab": []}
        grp, gi = None, 0
        for k in range(n):
            s = _text(els[k]).strip()
            nxt = els[k + 1] if k + 1 < n else None
            if s.startswith(("Đại diện TVQLDA", "Đại diện TVGS")):
                grp, gi = "tv", 0
            elif s.startswith(("Đại diện nhà thầu thi công", "Nhà thầu thi công xây dựng")):
                grp, gi = "nt", 0
            elif s.startswith(("Phòng thí nghiệm", "3. Thời gian", "4. Thời gian", "Đại diện chủ đầu tư")):
                grp = None
            elif grp and re.match(r"^(- ?)?Ông", s):
                pl = people[grp]
                if gi < len(pl) and pl[gi][0]:
                    pre = "- " if s.startswith("-") else ""
                    _set_text(els[k], f"{pre}Ông : {pl[gi][0]}          Chức vụ : {pl[gi][1]}")
                gi += 1
                continue
            s2 = s.lstrip("- ").strip()
            if "(ghi rõ tên biên bản kiểm tra)" in s:  # bỏ chữ gợi ý của form, thay bằng dấu chấm
                _set_text(els[k], s.replace("(ghi rõ tên biên bản kiểm tra)", "……………"))
                continue
            if s.startswith("Quảng Ninh, ngày"):
                _set_text(els[k], f"Quảng Ninh, {ngay}")  # phiếu yêu cầu: tạm lấy ngày nghiệm thu, người lập sửa nếu lập trước
            elif s2.startswith("Phiếu yêu cầu nghiệm thu của nhà thầu"):
                so_yc = rec.get("so_yc") or f"{_num(rec)}/{SUFFIX.get('YC' + t, 'YC/CV') if t in ('VL', 'CV', 'GD') else 'YC/CV'}"
                _set_text(els[k], f"- Phiếu yêu cầu nghiệm thu của nhà thầu thi công xây dựng số: {so_yc}")
            elif s.startswith("Phòng thí nghiệm") and project.get("phong_thi_nghiem"):
                _set_text(els[k], f"Phòng thí nghiệm: {project['phong_thi_nghiem']}")
            elif s2.startswith("Biên bản nghiệm thu công việc xây dựng liên quan") and rec.get("cv_ref"):
                _set_text(els[k], "- Biên bản nghiệm thu công việc xây dựng liên quan số: " + ", ".join(rec["cv_ref"]))
            elif s2.startswith("Bảng tính chi tiết khối lượng") and rec.get("bang_kl"):
                _set_text(els[k], f"- Bảng tính chi tiết khối lượng công việc thực hiện: {rec['bang_kl']}")
            elif s2.startswith("Tiêu chuẩn, quy phạm xây dựng") and rec.get("tieu_chuan"):
                _set_text(els[k], "- Tiêu chuẩn, quy phạm xây dựng được áp dụng: " + "; ".join(rec["tieu_chuan"]) + "; hồ sơ thiết kế, chỉ dẫn kỹ thuật của hợp đồng")
            elif s2.startswith("Tiêu chuẩn, quy phạm xây dựng"):  # chưa có tiêu chuẩn: bỏ chữ gợi ý, để dấu chấm
                _set_text(els[k], "- Tiêu chuẩn, quy phạm xây dựng được áp dụng: ………………………………………………………")
            elif s.startswith("5. Tiêu chuẩn kỹ thuật áp dụng") and nxt is not None and rec.get("tieu_chuan"):
                _set_text(nxt, "; ".join(rec["tieu_chuan"]))
            elif s2.startswith("Hồ sơ mời thầu") and project.get("so_hop_dong"):
                _set_text(els[k], f"- Hồ sơ mời thầu, hồ sơ trúng thầu và hợp đồng xây dựng số: {project['so_hop_dong']}")
            elif s2.startswith("Biên bản lấy mẫu, giao nhận mẫu") and rec.get("lm_ref"):
                _set_text(els[k], f"- Biên bản lấy mẫu, giao nhận mẫu thí nghiệm số: {', '.join(rec['lm_ref'])}")
            elif s2.startswith("Biên bản nghiệm thu vật liệu, thiết bị") and rec.get("vl_ref"):
                _set_text(els[k], f"- Biên bản nghiệm thu vật liệu, thiết bị, sản phẩm chế tạo sẵn trước khi sử dụng số: {', '.join(rec['vl_ref'])}")
            elif s2.startswith("Biên bản kiểm tra công tác") and rec.get("so_kt"):
                _set_text(els[k], f"- Biên bản kiểm tra công tác: Checklist kiểm tra, số {rec['so_kt']}")
            elif re.match(r"^Tại\s*(dự án)?\s*:?\s*(Dự án)?\s*[….]*$", s) and project.get("ten_cong_trinh") and t != "LM":
                _set_text(els[k], f"Tại dự án: {project['ten_cong_trinh']}")
            elif s.startswith("6. Khối lượng, quy cách lấy mẫu") and nxt is not None and rec.get("lm_quy_cach"):
                _set_text(nxt, rec["lm_quy_cach"])
            elif s.startswith("6. Khối lượng, quy cách lấy mẫu") and nxt is not None and rec.get("khoi_luong"):
                _set_text(nxt, f"Lô vật liệu khối lượng {_kl(rec['khoi_luong'])} {rec.get('dvt', '')}. Số lượng, quy cách mẫu: ……………………")
            elif s.startswith(("Biên bản số", "Phiếu yêu cầu số")):
                _set_text(els[k], f"{prefix} {_num(rec)}/{SUFFIX[t]}")
            elif s.startswith("Dự án:") and project.get("ten_cong_trinh"):
                _set_text(els[k], f"Dự án: {project['ten_cong_trinh']}")
            elif s.startswith("Địa điểm"):  # không giữ địa điểm của dự án mẫu
                _set_text(els[k], f"Địa điểm       :  {project.get('dia_diem') or '……………………………………………'}")
            elif s.startswith("Gói thầu:") and project.get("goi_thau"):
                _set_text(els[k], f"Gói thầu: {project['goi_thau']}")
            elif s.startswith(("Đại diện nhà thầu thi công xây dựng công trình", "Nhà thầu thi công xây dựng công trình")) \
                    and project.get("nha_thau"):
                _set_text(els[k], f"{s.split(':')[0]}: {project['nha_thau']}")
            elif s.startswith("Nhà thầu thi công ....") and project.get("nha_thau"):
                _set_text(els[k], re.sub(r"\.{5,}", project["nha_thau"], s, count=1))
            elif s.startswith("1. Tên công việc nghiệm thu") and nxt is not None:
                _set_text(nxt, rec.get("cong_viec") or f"Nghiệm thu {rec['hang_muc']}")
            elif s.startswith("2. Đối tượng kiểm tra"):
                _set_text(els[k], f"2. Đối tượng kiểm tra : {rec.get('lm_doi_tuong') or _doi_tuong(rec)}")
            elif s.startswith("1. Đối tượng nghiệm thu"):
                if s.rstrip().endswith(":") and nxt is not None:  # phiếu yêu cầu: nội dung ở đoạn kế
                    _set_text(nxt, _doi_tuong(rec))
                else:
                    _set_text(els[k], f"1. Đối tượng nghiệm thu : {_doi_tuong(rec)}")
            elif s.startswith(("Bắt đầu", "Kết thúc")):
                label = "Bắt đầu  :" if s.startswith("Bắt đầu") else "Kết thúc :"
                ngay_ = ngay
                if s.startswith("Kết thúc") and rec.get("ngay_den"):  # lấy mẫu bê tông kéo dài theo ngày đổ
                    d2, m2, y2 = _d(rec["ngay_den"])
                    ngay_ = f"ngày {d2:02d} tháng {m2:02d} năm {y2}"
                _set_text(els[k], f"{label}   {ngay_}")

        # Kết luận: chỉ điền khi người kiểm tra đã nhập.
        kl = rec.get("ket_luan", "")
        if t in ("VL", "CV") and kl and kl != "Chưa có kết luận":
            for k in range(n - 1):
                if _text(els[k]).strip().startswith("5. Kết luận"):
                    _set_text(els[k + 1], kl)

        _ten_nguoi_ky(els, doc, t, project)
        for el in els:
            _chuan_font(el)

        # Bảng khối lượng của form VL.
        if t == "VL":
            for el in els:
                if el.tag == qn('w:tbl'):
                    tb = Table(el, doc)
                    if tb.rows[0].cells[0].text.strip() == "STT" and len(tb.rows) > 1:
                        vals = ["1", rec.get("hang_muc", ""), "", rec.get("dvt", ""),
                                _kl(rec["khoi_luong"]) if rec.get("khoi_luong") else ""]
                        for c, v in zip(tb.rows[1].cells, vals):
                            if v:
                                c.paragraphs[0].add_run(v)
            for e2 in els:
                _chuan_font(e2)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def expand(records, kinds=("bb",)):
    """Từ biên bản VL/CV sinh danh sách biểu mẫu cần xuất.

    kinds: "bb" biên bản nghiệm thu, "yc" phiếu yêu cầu nghiệm thu, "lm" biên bản lấy mẫu (chỉ VL có lay_mau).
    Phiếu yêu cầu dùng cùng số với biên bản; biên bản lấy mẫu đánh số riêng theo thứ tự.
    """
    out, lm_no = [], 0
    for r in sorted(records, key=lambda r: (r["so"].split()[0], int(re.sub(r"\D", "", r["so"].split()[-1]) or 0))):  # theo số hiệu
        if r["type"] not in ("VL", "CV", "GD"):
            continue
        if "yc" in kinds:
            out.append(dict(r, type={"VL": "YCVL", "CV": "YCCV", "GD": "YCGD"}[r["type"]]))
        if "lm" in kinds and r.get("lay_mau") and r["type"] in ("VL", "CV"):
            lm_no += 1
            cv_ = (r.get("lm_cong_viec") or f"Lấy mẫu thí nghiệm phục vụ nghiệm thu vật liệu đầu vào: {r['hang_muc']}" if r["type"] == "VL"
                   else f"Lấy mẫu bê tông thí nghiệm R7, R28 phục vụ nghiệm thu công việc: {r['hang_muc']}")
            out.append(dict(r, type="LM", so=r.get("so_lm") or f"LM {lm_no:03d}", ngay=r.get("ngay_lay_mau") or r["ngay"], ngay_den=r.get("ngay_lay_mau_den"),
                            cong_viec=cv_, tieu_chuan=r.get("lm_tieu_chuan") or r.get("tieu_chuan")))
        if "bb" in kinds:
            out.append(r)
    if set(kinds) == {"lm"}:
        out.sort(key=lambda r: r["so"])
    return out
