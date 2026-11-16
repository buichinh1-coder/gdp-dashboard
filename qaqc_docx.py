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
}
SUFFIX = {"VL": "CV/VL-TB", "CV": "CV", "LM": "LM", "YCVL": "YC/VL-TB", "YCCV": "YC/CV"}
IS_YC = ("YCVL", "YCCV")


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
    if rec["type"] in ("CV", "YCCV", "YCVL") and rec.get("khoi_luong"):
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


def _ten_nguoi_ky(els, doc, t, project):
    tung, anh = project.get("chi_huy_truong"), project.get("can_bo_tvgs")
    for el in els:
        if el.tag != qn('w:tbl'):
            continue
        tb = Table(el, doc)
        head = tb.rows[0].cells[0].text.strip()
        if t == "LM" and head.startswith("Đại diện TVGS"):
            if anh:
                _name_in_cell(tb.rows[1].cells[0], anh, 4)
            if tung:
                _name_in_cell(tb.rows[1].cells[1], tung, 4)
        elif t == "CV" and head.startswith("Đại diện TVGS"):
            c1 = tb.rows[0].cells[1]  # bớt 1 dòng trống đầu ô tiêu đề (khớp file mẫu bản chỉnh của người dùng)
            if len(c1.paragraphs) >= 7 and not c1.paragraphs[0].text.strip() and not c1.paragraphs[1].text.strip():
                c1._tc.remove(c1.paragraphs[0]._p)
            if anh:
                _name_in_cell(tb.rows[1].cells[0], anh, 2)
            if tung:
                _name_in_cell(tb.rows[1].cells[1], tung, 2)
        elif t in IS_YC and head.startswith("Đơn vị TVQLDA"):
            c0, c1 = tb.rows[1].cells
            if anh:
                for p in c0.paragraphs[1:]:
                    for tt in p._p.findall('.//' + qn('w:t')):
                        tt.text = ''
                _set_text(c0.paragraphs[0]._p, f"Người nhận: {anh}")
            if tung:
                _set_text(c1.paragraphs[-1]._p, tung)
        elif t == "VL" and head.startswith("Đại diện TVQLDA"):
            if anh and tung:
                cell = tb.rows[1].cells[0]
                ps = cell.paragraphs
                _r = ps[3].add_run()
                _r.add_text(" " * 7 + anh + " " * 47 + tung)
                _r.bold, _r.font.name = True, "Times New Roman"


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
            elif s.startswith(("Phòng thí nghiệm", "3. Thời gian", "4. Thời gian")):
                grp = None
            elif grp and re.match(r"^(- ?)?Ông", s):
                pl = people[grp]
                if gi < len(pl) and pl[gi][0]:
                    pre = "- " if s.startswith("-") else ""
                    _set_text(els[k], f"{pre}Ông : {pl[gi][0]}          Chức vụ : {pl[gi][1]}")
                gi += 1
                continue
            s2 = s.lstrip("- ").strip()
            if s.startswith("Quảng Ninh, ngày"):
                _set_text(els[k], f"Quảng Ninh, {ngay}")  # phiếu yêu cầu: tạm lấy ngày nghiệm thu, người lập sửa nếu lập trước
            elif s2.startswith("Phiếu yêu cầu nghiệm thu của nhà thầu"):
                so_yc = rec.get("so_yc") or f"{_num(rec)}/{'YC/VL-TB' if t == 'VL' else 'YC/CV'}"
                _set_text(els[k], f"- Phiếu yêu cầu nghiệm thu của nhà thầu thi công xây dựng số: {so_yc}")
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
                _set_text(els[k], f"2. Đối tượng kiểm tra : {_doi_tuong(rec)}")
            elif s.startswith("1. Đối tượng nghiệm thu"):
                if s.rstrip().endswith(":") and nxt is not None:  # phiếu yêu cầu: nội dung ở đoạn kế
                    _set_text(nxt, _doi_tuong(rec))
                else:
                    _set_text(els[k], f"1. Đối tượng nghiệm thu : {_doi_tuong(rec)}")
            elif s.startswith(("Bắt đầu", "Kết thúc")):
                label = "Bắt đầu  :" if s.startswith("Bắt đầu") else "Kết thúc :"
                _set_text(els[k], f"{label}   {ngay}")

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
    for r in sorted(records, key=lambda r: (r["ngay"], r["so"])):
        if r["type"] not in ("VL", "CV"):
            continue
        if "yc" in kinds:
            out.append(dict(r, type="YCVL" if r["type"] == "VL" else "YCCV"))
        if "lm" in kinds and r["type"] == "VL" and r.get("lay_mau"):
            lm_no += 1
            out.append(dict(r, type="LM", so=r.get("so_lm") or f"LM {lm_no:03d}", cong_viec=f"Lấy mẫu thí nghiệm phục vụ nghiệm thu vật liệu đầu vào: {r['hang_muc']}"))
        if "bb" in kinds:
            out.append(r)
    return out
