"""Thuyết minh biện pháp thi công (Word). Số liệu lấy từ lịch nghiệm thu và records.json; không tự đặt số liệu thiết kế.
Dùng: python tools/build_bptc_thuyet_minh.py <lich.xlsx> <file_ra.docx>
"""
import datetime as dt
import json
import sys
from pathlib import Path

import docx
import openpyxl
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).parent.parent
PRJ = json.loads((ROOT / "data/qaqc/records.json").read_text(encoding="utf-8"))["project"]
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]
dmy = lambda d: d.strftime("%d/%m/%Y")


def vn(x, nd=2):
    return f"{x:,.{nd}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Chi tiet nghiem thu"]
    out = []
    for r in ws.iter_rows(min_row=9, values_only=True):
        if r[0] is None or r[6] is None:
            continue
        muc = str(r[1] or "")
        if muc == "16":
            muc = "16b" if str(r[2]).lower().startswith("hoàn thiện") else "16a"
        out.append(dict(muc=muc, pd=r[3], dvt=r[4], kl=r[5] or 0, s=r[6].date(), e=r[7].date()))
    return out


def main(lich, out):
    data = load(lich)
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(1.5), Cm(1.5), Cm(1.5)
    st = d.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(13)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.15

    def P(text="", bold=False, italic=False, align=None, size=None, indent=None, after=None, keep=False):
        p = d.add_paragraph()
        if align == "c":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "j":
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent:
            p.paragraph_format.left_indent = Cm(indent)
        if after is not None:
            p.paragraph_format.space_after = Pt(after)
        if keep:
            p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        if size:
            r.font.size = Pt(size)
        return p

    def H1(t):
        p = P(t, bold=True, size=14, after=6, keep=True)
        p.paragraph_format.space_before = Pt(10)
        return p

    def H2(t):
        p = P(t, bold=True, after=3, keep=True)
        p.paragraph_format.space_before = Pt(6)
        return p

    def B(t, lvl=0):
        p = d.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9 + 0.7 * lvl)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run("- " + t)
        return p

    def T(head, rows, widths=None, size=12, center_cols=()):
        t = d.add_table(rows=1, cols=len(head))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(head):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(h)
            r.bold, r.font.size = True, Pt(size)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            sh = OxmlElement("w:shd")
            sh.set(qn("w:val"), "clear"), sh.set(qn("w:fill"), "D9D9D9")
            c._tc.get_or_add_tcPr().append(sh)
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = ""
                r = cells[i].paragraphs[0].add_run(str(v))
                r.font.size = Pt(size)
                if i in center_cols:
                    cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.autofit = False
        if widths:
            for i, w in enumerate(widths):
                t.columns[i].width = Cm(w)
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        # lặp dòng tiêu đề
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
        d.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    # ---------------- chân trang: số trang
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("Thuyết minh biện pháp thi công - " + PRJ["nha_thau"] + " - Trang ")
    r.font.size = Pt(10)
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        rr = fp.add_run()
        rr.font.size = Pt(10)
        if tag:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), tag)
            rr._r.append(f)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            rr._r.append(it)

    # ---------------- bìa
    for _ in range(2):
        P("")
    P("CÔNG TY TNHH GIANG LONG", bold=True, align="c", size=15)
    P("")
    P("THUYẾT MINH BIỆN PHÁP THI CÔNG", bold=True, align="c", size=22, after=8)
    P("GÓI THẦU: " + PRJ["goi_thau"].upper(), bold=True, align="c", size=15, after=14)
    P("Địa điểm: " + PRJ["dia_diem"], align="c")
    P("Chủ đầu tư: " + PRJ["chu_dau_tu"], align="c")
    P("Tư vấn quản lý dự án và giám sát: Công ty TNHH Apave Châu Á – Thái Bình Dương", align="c")
    P("Nhà thầu thi công: " + PRJ["nha_thau"], align="c")
    P("Hợp đồng số " + PRJ["so_hop_dong"], align="c")
    P("Thời gian thi công: 12/08/2026 – 27/10/2026", align="c", after=18)
    P("Kèm theo: Bộ bản vẽ biện pháp thi công BV-01 đến BV-06 (A3) và Tiến độ thi công.", italic=True, align="c", after=40)
    tb = d.add_table(rows=2, cols=3)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (a, b) in enumerate([("ĐẠI DIỆN NHÀ THẦU THI CÔNG", PRJ["nha_thau"].upper()), ("TƯ VẤN QUẢN LÝ DỰ ÁN VÀ GIÁM SÁT", "CÔNG TY TNHH APAVE CHÂU Á – THÁI BÌNH DƯƠNG"),
                                 ("CHỦ ĐẦU TƯ CHẤP THUẬN", PRJ["chu_dau_tu"].upper())]):
        c = tb.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.paragraphs[0].add_run(a + "\n")
        r.bold, r.font.size = True, Pt(12)
        r2 = c.paragraphs[0].add_run(b)
        r2.bold, r2.font.size = True, Pt(11)
        c2 = tb.rows[1].cells[i]
        c2.text = ""
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = c2.paragraphs[0].add_run("(Ký, ghi rõ họ tên, đóng dấu)")
        r3.italic, r3.font.size = True, Pt(11)
        c2.add_paragraph("\n\n\n\n")
    d.add_page_break()

    # ---------------- mục lục
    P("MỤC LỤC", bold=True, align="c", size=14, after=8)
    toc = ["1. Cơ sở lập biện pháp thi công", "2. Giới thiệu công trình và điều kiện thi công", "3. Tổ chức thi công", "4. Mặt bằng thi công và công tác chuẩn bị",
           "5. Trình tự và tiến độ thi công tổng thể", "6. Biện pháp thi công chi tiết", "7. Biện pháp quản lý chất lượng và nghiệm thu", "8. Biện pháp an toàn lao động",
           "9. Vệ sinh môi trường, phòng cháy chữa cháy, an ninh", "10. Biện pháp ứng phó rủi ro", "11. Kết luận", "Phụ lục: Danh mục bản vẽ và tài liệu kèm theo"]
    for t in toc:
        P(t, after=2)
    d.add_page_break()

    # ---------------- 1
    H1("1. CƠ SỞ LẬP BIỆN PHÁP THI CÔNG")
    for t in [f"Hợp đồng số {PRJ['so_hop_dong']} giữa {PRJ['chu_dau_tu']} và {PRJ['nha_thau']} về gói thầu \"{PRJ['goi_thau']}\".",
              "Hồ sơ thiết kế, bản vẽ thi công, chỉ dẫn kỹ thuật và hồ sơ cấp phối bê tông (đã được duyệt) của gói thầu. Các số liệu kích thước, cao độ, quy cách vật liệu, lực ép, độ chối, chiều dày lớp đắp trong tài liệu này lấy theo hồ sơ đó.",
              "Lịch nghiệm thu và tiến độ thi công đã thống nhất (12/08/2026 – 27/10/2026, làm 7 ngày/tuần, nghỉ lễ 01–02/09).",
              "Điều kiện hiện trường và mặt bằng do Chủ đầu tư bàn giao."]:
        B(t)
    P("Các tiêu chuẩn, quy chuẩn áp dụng (xác nhận hiệu lực với Chủ đầu tư và tư vấn giám sát tại thời điểm thi công):", after=2)
    for t in ["TCVN 9394:2012 – Đóng và ép cọc, thi công và nghiệm thu.", "TCVN 4447:2012 – Công tác đất, thi công và nghiệm thu.",
              "TCVN 9361:2012 – Công tác nền móng, thi công và nghiệm thu.", "TCVN 4453:1995 – Kết cấu bê tông và bê tông cốt thép toàn khối, quy phạm thi công và nghiệm thu (cần xác nhận hiệu lực/tiêu chuẩn thay thế).",
              "TCVN 3105:2022 – Hỗn hợp bê tông và bê tông, lấy mẫu, chế tạo và bảo dưỡng mẫu thử; TCVN 3118:2022 – Xác định cường độ chịu nén.",
              "TCVN 8828:2011 – Bê tông, yêu cầu bảo dưỡng ẩm tự nhiên.", "TCVN 1651-2:2018 – Thép cốt bê tông, thép thanh vằn; QCVN 07:2019/BKHCN – thép làm cốt bê tông.",
              "TCVN 7570:2006 – Cốt liệu cho bê tông và vữa; TCVN 4506:2012 – Nước cho bê tông và vữa.",
              "Các quy định hiện hành về quản lý chất lượng, an toàn lao động và vệ sinh môi trường trong thi công xây dựng."]:
        B(t)

    # ---------------- 2
    H1("2. GIỚI THIỆU CÔNG TRÌNH VÀ ĐIỀU KIỆN THI CÔNG")
    T(["Nội dung", "Thông tin"], [["Gói thầu", PRJ["goi_thau"]], ["Địa điểm", PRJ["dia_diem"]], ["Chủ đầu tư", PRJ["chu_dau_tu"]],
                                  ["Tư vấn QLDA và giám sát", "Công ty TNHH Apave Châu Á – Thái Bình Dương"], ["Nhà thầu thi công", PRJ["nha_thau"]],
                                  ["Chỉ huy trưởng công trường", PRJ["chi_huy_truong"]], ["Thời gian thi công", "12/08/2026 – 27/10/2026 (77 ngày, làm 7 ngày/tuần)"],
                                  ["Phòng thí nghiệm", PRJ["phong_thi_nghiem"]]], [5, 12])
    P("Công trình là tuyến kè chắn đất bê tông cốt thép đổ tại chỗ phía núi, tuyến 3, đoạn từ cọc 88 đến cọc 126 (chiều dài tuyến theo lịch: 449,16 m), chia làm 4 phân đoạn thi công cuốn chiếu: cọc 88–98, cọc 98–108, cọc 108–118 và cọc 118–126. Kết cấu gồm cọc bê tông cốt thép 300x300 (M300 bền sunfat), bê tông lót M100, móng - thân - khóa đỉnh bê tông M300 bền sunfat có phụ gia chống thấm B12, khe lún, hoàn trả hố móng phía trong kè, đá xô chân tường chắn phía ngoài kè.", align="j")
    P("Khối lượng chính theo lịch nghiệm thu (khối lượng từng phân đoạn là số phân bổ tạm, cập nhật theo bản vẽ triển khai):", after=2)
    NAMES = [("15", "Đắp đê quai, sàn thi công, thoát nước tạm"), ("16a", "Đào bùn, nạo vét lạch phía núi (đợt mở mặt bằng)"), ("01", "Cọc BT300x300 M300 R7 (m)"),
             ("04", "Đào móng bằng máy, đất cấp II"), ("02", "Ép cọc (m)"), ("10", "Ván khuôn bê tông lót (m2)"), ("08", "Bê tông lót M100"),
             ("13", "Cốt thép móng, thân, khóa đỉnh (tấn)"), ("11", "Ván khuôn móng, thân (m2)"), ("12", "Ván khuôn khóa đỉnh (m2)"), ("14", "Khe lún, chi tiết thoát nước (m2)"),
             ("09", "Bê tông M300 bền sunfat B12"), ("06", "Đắp hoàn trả hố móng phía trong kè"), ("07", "Đắp đá xô chân tường chắn"), ("16b", "Hoàn thiện nạo vét (30% còn lại)")]
    rows = []
    for k, n in NAMES:
        vals = []
        for p in PD:
            v = sum(x["kl"] for x in data if x["muc"] == k and x["pd"] == p)
            vals.append(vn(v))
        tot = sum(x["kl"] for x in data if x["muc"] == k)
        dv = next(x["dvt"] for x in data if x["muc"] == k)
        rows.append([n, dv] + vals + [vn(tot)])
    T(["Hạng mục", "ĐVT", "PĐ1", "PĐ2", "PĐ3", "PĐ4", "Tổng"], rows, [5.6, 1.4, 2.0, 2.0, 2.0, 2.0, 2.2], 10.5, (1, 2, 3, 4, 5, 6))
    P("Đặc điểm và điều kiện thi công cần lưu ý: (1) tuyến kè ở phía núi, có lạch phía núi cần nạo vét để mở mặt bằng; (2) hố móng nằm gần nguồn nước nên cần đê quai, sàn thi công và hệ thống thoát nước, bơm; (3) thi công trong mùa mưa có thể xảy ra, cần phương án che chắn và dự phòng; (4) mặt bằng hẹp theo tuyến, xe máy thi công di chuyển cuốn chiếu theo phân đoạn. Các nhận định này cần đối chiếu với khảo sát hiện trường và hồ sơ địa chất của dự án.", align="j")

    # ---------------- 3
    H1("3. TỔ CHỨC THI CÔNG")
    H2("3.1. Sơ đồ tổ chức hiện trường")
    T(["Vị trí", "Họ tên", "Nhiệm vụ chính"],
      [["Chỉ huy trưởng công trường", PRJ["chi_huy_truong"], "Điều hành chung, chịu trách nhiệm tiến độ, chất lượng, an toàn; ký hồ sơ nghiệm thu"],
       ["Cán bộ kỹ thuật thi công", "(nhà thầu bổ sung)", "Triển khai biện pháp, đo đạc, lập hồ sơ, nghiệm thu nội bộ"],
       ["Cán bộ QA/QC, lấy mẫu", "(nhà thầu bổ sung)", "Kiểm soát vật liệu, checklist, lấy mẫu thí nghiệm"],
       ["Cán bộ an toàn - môi trường", "(nhà thầu bổ sung)", "Kiểm tra an toàn, vệ sinh môi trường hằng ngày"],
       ["Đội trưởng các đội thi công", "(nhà thầu bổ sung)", "Đội cọc, đội đất, đội cốt thép - ván khuôn, đội bê tông"],
       ["Cán bộ tư vấn giám sát (Apave)", PRJ["can_bo_tvgs"], "Kiểm tra, nghiệm thu từng công việc"],
       ["Giám sát trưởng (Apave)", PRJ["giam_sat_truong"], "Chỉ đạo giám sát, chấp thuận hồ sơ"]], [5, 4.3, 8.5])
    H2("3.2. Máy móc, thiết bị chủ yếu (dự kiến, nhà thầu xác nhận số lượng và chủng loại)")
    T(["Thiết bị", "Công việc sử dụng", "Số lượng"],
      [["Máy đào, máy ủi", "Đắp đê quai, nạo vét, đào móng, san đắp", "…"], ["Ô tô tự đổ", "Vận chuyển đất đào, đất đắp, đá xô trong phạm vi quy định", "…"],
       ["Máy ép cọc (đủ lực ép theo thiết kế) và cẩu", "Ép cọc BTCT 300x300", "…"], ["Máy cắt/đục bê tông, búa phá", "Cắt, đập đầu cọc", "…"],
       ["Máy cắt, uốn thép; máy hàn", "Gia công cốt thép, nối cọc (nếu có)", "…"], ["Bơm bê tông hoặc xe bồn, đầm dùi", "Đổ, đầm bê tông M300 và M100", "…"],
       ["Máy bơm nước", "Thoát nước hố móng, nước mưa", "…"], ["Lu rung, đầm cóc", "Đầm nén đất đắp K95", "…"],
       ["Máy toàn đạc/thủy bình", "Định vị tim, cao độ", "…"], ["Máy phát điện, tủ điện, chiếu sáng", "Cấp điện, làm việc khi cần", "…"]], [6.2, 8, 3.2], 12, (2,))
    H2("3.3. Nhân lực")
    P("Số lượng công nhân từng đội do nhà thầu bố trí theo tiến độ thi công cuốn chiếu; tại mỗi thời điểm các đội tập trung vào một đến hai phân đoạn. Công nhân phải được huấn luyện an toàn và trang bị bảo hộ lao động trước khi vào công trường.", align="j")

    # ---------------- 4
    H1("4. MẶT BẰNG THI CÔNG VÀ CÔNG TÁC CHUẨN BỊ")
    P("Bố trí mặt bằng thi công thể hiện trong bản vẽ BV-01. Vị trí cụ thể các khu phụ trợ theo hiện trạng mặt bằng được Chủ đầu tư bàn giao.", align="j")
    for t in ["Nhận bàn giao mặt bằng, tim mốc, cao độ chuẩn; lập biên bản; đối chiếu với bản vẽ thiết kế; bảo vệ mốc trong suốt quá trình thi công.",
              "Đo đạc định vị tim tuyến kè, tim cọc, cao độ hố móng bằng máy toàn đạc/thủy bình; kết quả đo kiểm tra lại trước mỗi công việc quan trọng.",
              "Làm đường công vụ và sàn thi công cho xe máy di chuyển, giữ thông thoát nước tạm, bố trí rào chắn và biển báo.",
              "Dựng văn phòng chỉ huy, kho vật tư, bãi gia công cốt thép - ván khuôn, bãi tập kết cọc BTCT và bãi tạm đất, đá xô.",
              "Cấp điện, cấp nước thi công, hệ thống bơm thoát nước, hố lắng, khu rửa xe ra vào.",
              "Trình hồ sơ vật liệu (cọc, thép, cấp phối bê tông, đá xô, đất đắp) để tư vấn giám sát chấp thuận trước khi đưa vào sử dụng."]:
        B(t)

    # ---------------- 5
    H1("5. TRÌNH TỰ VÀ TIẾN ĐỘ THI CÔNG TỔNG THỂ")
    P("Thi công theo phương pháp cuốn chiếu qua bốn phân đoạn, mỗi phân đoạn thực hiện theo trình tự sau (xem BV-03):", after=2)
    for i, t in enumerate(["Đắp đê quai, tạo sàn thi công và tuyến thoát nước tạm.", "Đào bùn, nạo vét lạch phía núi (đợt mở mặt bằng).",
                           "Cung cấp và nghiệm thu đầu vào cọc BTCT 300x300 (đạt R7), cốt thép.", "Đào móng bằng máy theo từng đốt, vận chuyển đất đào.",
                           "Ép cọc, cắt đập đầu cọc.", "Ván khuôn và đổ bê tông lót M100.", "Cốt thép, ván khuôn móng/thân/khóa đỉnh, khe lún và chi tiết thoát nước.",
                           "Đổ bê tông M300 bền sunfat, lấy mẫu R7/R28, bảo dưỡng.", "Đắp hoàn trả hố móng phía trong kè theo lớp, đầm đạt K95.",
                           "Đắp đá xô chân tường chắn phía ngoài kè.", "Hoàn thiện nạo vét lạch phía núi (30% khối lượng còn lại)."], 1):
        B(f"Bước {i}: {t}")
    P("Thời gian thi công từng phân đoạn theo lịch nghiệm thu:", after=2)
    rows = []
    for p in PD:
        ds = [x for x in data if x["pd"] == p]
        rows.append([p.replace("Cọc", "Cọc"), dmy(min(x["s"] for x in ds)), dmy(max(x["e"] for x in ds)),
                     dmy(next(x["e"] for x in ds if x["muc"] == "02")) if False else f"{dmy(min(x['s'] for x in ds if x['muc'] == '09'))} – {dmy(max(x['e'] for x in ds if x['muc'] == '09'))}"])
    T(["Phân đoạn", "Bắt đầu", "Kết thúc (hoàn thiện)", "Đổ bê tông M300"], rows, [4, 3.2, 4.5, 5.2], 12, (1, 2, 3))
    P("Công việc cuối cùng của cả gói là nghiệm thu hoàn thành (26–27/10/2026) và bàn giao công trình, hồ sơ (27/10/2026). Đường găng của tiến độ đi theo chuỗi công việc của phân đoạn 4 (cọc 118–126), nên phân đoạn này cần ưu tiên nhân lực và thiết bị (chi tiết trong bản Tiến độ thi công đính kèm).", align="j")

    # ---------------- 6
    H1("6. BIỆN PHÁP THI CÔNG CHI TIẾT")

    def block(title, items):
        H2(title)
        for lab, tx in items:
            p = d.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r = p.add_run("- " + lab + ": ")
            r.bold = True
            p.add_run(tx)

    block("6.1. Đắp đê quai, sàn thi công và thoát nước tạm (BV-01, BV-02)",
          [("Biện pháp", "Đắp đê quai dọc phía lạch và phía hố móng bằng đất đủ chất lượng, đầm theo lớp, tạo sàn cho xe máy thi công; làm rãnh thoát nước tạm và hố bơm ở điểm thấp."),
           ("Yêu cầu", "Cao độ, mặt cắt đê quai theo thiết kế; mái đê ổn định, không thấm gây ngập hố móng. Độ chặt đắp đạt K95 theo yêu cầu của dự án."),
           ("Kiểm soát", "Kiểm tra độ chặt tại hiện trường theo từng ngày đắp; kiểm tra cao độ, kích thước; nghiệm thu công việc trước khi chuyển bước."),
           ("An toàn", "Rào chắn, biển báo; không để thiết bị nặng sát mép; theo dõi mực nước trước và sau mưa.")])
    block("6.2. Đào bùn, nạo vét lạch phía núi (BV-01, BV-02)",
          [("Biện pháp", "Nạo vét bằng máy đào, chia hai đợt: đợt mở mặt bằng (70% khối lượng theo lịch) và đợt hoàn thiện (30% còn lại sau khi hoàn trả và đắp đá xô). Bùn đất quăng tải trong phạm vi quy định, vận chuyển đến nơi đổ được chấp thuận."),
           ("Kiểm soát", "Kiểm tra cao độ, mặt cắt lạch bằng đo đạc; lập biên bản nghiệm thu từng đợt."),
           ("Môi trường", "Hố lắng, không xả bùn trực tiếp ra nguồn nước; che phủ xe chở bùn.")])
    block("6.3. Cọc bê tông cốt thép 300x300: vật liệu đầu vào, ép cọc, cắt đầu cọc (BV-04)",
          [("Vật liệu đầu vào", "Cọc 300x300 M300 bền sunfat nhà thầu tự sản xuất; nghiệm thu đầu vào một lần cho cả lô khi đủ R7 (phụ gia R7 đạt cường độ thiết kế sau 7 ngày) và hồ sơ được chấp thuận; thép chế tạo cọc có chứng chỉ và kết quả thí nghiệm kéo; lấy mẫu bê tông cọc kiểm tra cường độ (biên bản lấy mẫu theo hồ sơ). Chỉ được ép sau khi cọc đạt R7 và được chấp thuận."),
           ("Chuẩn bị", "Định vị tim cọc theo bản vẽ; kiểm tra kích thước, độ thẳng, khuyết tật cọc trước khi cẩu; đánh dấu chiều dài trên thân cọc để theo dõi độ sâu."),
           ("Ép cọc", "Đặt máy ép đúng tim, căn thẳng đứng; ép liên tục, tốc độ ép đều; ghi nhật ký ép từng cọc (lực ép, chiều dài, thời gian). Dừng ép khi đạt lực ép và độ chối theo thiết kế. Nối cọc (nếu thiết kế quy định) bằng hàn, đảm bảo thẳng trục. Thứ tự ép theo một hướng, tránh xô lệch cọc đã ép."),
           ("Cắt, đập đầu cọc", "Sau khi hoàn thành ép của phân đoạn và nghiệm thu, đập bỏ phần đầu cọc đến cao độ thiết kế, giữ nguyên cốt thép chờ để neo vào móng, thu gom bốc xúc và vận chuyển đầu cọc ra khỏi hố móng."),
           ("Kiểm soát", "Checklist ép cọc, cắt đầu cọc; biên bản nghiệm thu công việc; bảng tổng hợp thi công cọc; nghiệm thu giai đoạn thi công cọc toàn tuyến.")])
    block("6.4. Đào móng bằng máy và vận chuyển đất đào (BV-02)",
          [("Biện pháp", "Đào bằng máy theo từng đốt (đất cấp II) đến gần cao độ thiết kế, hoàn thiện đáy bằng nhân công; đào đến đâu thoát nước, hoàn thiện lớp lót ngay đến đó để tránh ngâm nước."),
           ("Yêu cầu", "Mái hố móng ổn định, đáy hố sạch, đúng cao độ; không đào quá sâu; bố trí rãnh thu nước và hố bơm."),
           ("Đất đào", "Quăng tải, xúc vận chuyển trong phạm vi quy định đến bãi tạm hoặc nơi đổ; phần đất đạt yêu cầu có thể dùng đắp hoàn trả nếu được tư vấn giám sát chấp thuận."),
           ("Kiểm soát", "Đo cao độ, kích thước đáy móng theo từng đốt; nghiệm thu hố móng trước khi ván khuôn lót và đổ bê tông lót.")])
    block("6.5. Ván khuôn và bê tông lót M100",
          [("Biện pháp", "Dựng ván khuôn lót theo tim và cao độ, đổ bê tông lót M100 thành lớp phẳng, đủ chiều dày theo thiết kế."),
           ("Lưu ý", "Bê tông lót không lấy mẫu thí nghiệm theo thỏa thuận của dự án; kiểm soát bằng nghiệm thu hình học và quan sát chất lượng.")])
    block("6.6. Gia công, lắp dựng cốt thép",
          [("Vật liệu", "Cốt thép nghiệm thu đầu vào một lần cho cả gói: chứng chỉ chất lượng, kết quả thí nghiệm kéo, kiểm tra mác thép, đường kính, dấu hiệu nhận biết theo TCVN 1651-2:2018 và QCVN 07:2019/BKHCN."),
           ("Gia công", "Cắt uốn đúng hình dạng, kích thước theo bản vẽ; mối nối buộc/hàn theo thiết kế; không dùng thép gỉ sâu, dính dầu mỡ, bùn."),
           ("Lắp dựng", "Đặt đúng vị trí, khoảng cách, lớp bê tông bảo vệ bằng con kê; neo cốt thép chờ đầu cọc vào móng đúng chiều dài thiết kế."),
           ("Nghiệm thu", "Nghiệm thu cốt thép, ván khuôn trước khi đổ bê tông móng, thân, khóa đỉnh theo từng đợt đổ.")])
    block("6.7. Ván khuôn móng, thân và khóa đỉnh (BV-05)",
          [("Biện pháp", "Ván khuôn định hình hoặc gỗ, ghép kín khít, có giằng, cây chống, thanh chống xiên đảm bảo cứng, ổn định dưới áp lực bê tông; bôi chất chống dính."),
           ("Kiểm tra", "Kích thước, tim, cao độ, độ thẳng đứng, độ kín khít, tính ổn định của hệ chống trước khi đổ bê tông."),
           ("Tháo dỡ", "Chỉ tháo ván khuôn thân khi bê tông đạt cường độ đủ để không hư hại cạnh góc và theo yêu cầu thiết kế/chỉ dẫn kỹ thuật.")])
    block("6.8. Đổ bê tông M300 bền sunfat có phụ gia chống thấm B12 (BV-05)",
          [("Vật liệu", "Bê tông thương phẩm theo cấp phối đã được duyệt; kiểm tra phiếu giao hàng, thời gian trộn - đổ, độ sụt theo phiếu; không cho thêm nước tại hiện trường."),
           ("Đổ và đầm", "Chia đợt đổ theo từng đốt (móng, thân, khóa đỉnh); đổ phân lớp, đầm dùi đủ thời gian, không chạm ván khuôn và cốt thép; mạch ngừng bố trí đúng thiết kế, xử lý trước khi đổ tiếp. Sử dụng xe bơm hoặc cần phân phối đưa bê tông tới vị trí, hạn chế rơi tự do."),
           ("Lấy mẫu", "Lấy mẫu tại hiện trường ngay ngày đổ theo từng đợt đổ, bảo dưỡng chuẩn và nén ở tuổi R7 và R28 (phụ gia R7 đạt cường độ thiết kế ở 7 ngày; R28 để nghiệm thu cường độ). Không làm thí nghiệm độ sụt tại hiện trường và thấm nước theo yêu cầu của chỉ dẫn kỹ thuật. Số tổ mẫu theo thể tích mỗi đợt đổ, tạm tính 1 tổ/20 m3 (mỗi tổ 3 viên), cần đối chiếu chỉ dẫn kỹ thuật và tư vấn giám sát; thí nghiệm do phòng thí nghiệm " + PRJ["phong_thi_nghiem"] + " thực hiện."),
           ("Bảo dưỡng", "Bảo dưỡng ẩm theo TCVN 8828:2011, che nắng mưa, tưới nước hoặc phủ bao tải ẩm; ghi nhật ký bảo dưỡng."),
           ("Nghiệm thu", "Biên bản nghiệm thu công việc đổ bê tông kèm biên bản lấy mẫu; kết quả R7, R28 bổ sung vào hồ sơ.")])
    block("6.9. Thi công khe lún và chi tiết thoát nước (BV-05)",
          [("Biện pháp", "Khe lún bố trí giữa các đốt theo bản vẽ, ván khuôn khe thẳng, vật liệu chèn khe và băng cản nước đặt trước khi đổ bê tông đốt kế tiếp; lắp ống/lỗ thoát nước theo đúng vị trí và độ dốc thiết kế."),
           ("Kiểm soát", "Nghiệm thu từng đợt cùng với ván khuôn - cốt thép, bảo đảm khe thẳng, không lọt vữa, không bịt thoát nước.")])
    block("6.10. Đắp hoàn trả hố móng phía trong kè, lu nền (BV-06)",
          [("Điều kiện", "Chỉ đắp khi kết cấu bê tông đạt cường độ cho phép chịu tải đắp theo thiết kế và đã nghiệm thu; dọn sạch bùn, rác, nước đọng trong hố."),
           ("Biện pháp", "Đất đắp đạt yêu cầu, đổ và san thành từng lớp có chiều dày theo chỉ dẫn kỹ thuật (xác định bằng đầm thử), đắp đối xứng, xe nặng không áp sát tường; lu/đầm đạt độ chặt K95."),
           ("Kiểm soát", "Kiểm tra độ chặt K95 tại hiện trường theo từng ngày đắp và từng lớp (phương pháp theo chỉ dẫn kỹ thuật); lớp trên chỉ đắp khi lớp dưới đạt.")])
    block("6.11. Đắp đá xô chân tường chắn phía ngoài kè (BV-06)",
          [("Vật liệu", "Đá xô nghiệm thu vật liệu đầu vào; đá sạch, không lẫn bùn đất, kích cỡ theo thiết kế."),
           ("Biện pháp", "Đổ/xếp đá chân tường tới cao độ và mặt cắt thiết kế, hoàn thiện mái gọn; không làm xê dịch kết cấu bê tông."),
           ("Kiểm soát", "Nghiệm thu công việc về hình học và chất lượng vật liệu.")])
    block("6.12. Hoàn thiện nạo vét lạch phía núi",
          [("Biện pháp", "Sau khi đắp hoàn trả và đá xô, nạo vét phần khối lượng còn lại (30%) hoàn thiện lòng lạch theo mặt cắt thiết kế, dọn sạch vật liệu thừa, thu dọn thiết bị tạm.")])
    block("6.13. Biện pháp thi công trọn gói: đường công vụ, đê quai, bơm và bảo vệ môi trường",
          [("Nội dung", "Duy trì đường công vụ cho thi công và dân sinh (nếu có), bảo trì đê quai, vận hành bơm thoát nước hố móng, hố lắng, che chắn, tưới nước giảm bụi, vệ sinh đường. Nghiệm thu tổng hợp tại bàn giao toàn gói.")])

    # ---------------- 7
    H1("7. BIỆN PHÁP QUẢN LÝ CHẤT LƯỢNG VÀ NGHIỆM THU")
    for t in ["Mọi công việc được kiểm tra bằng checklist, nghiệm thu nội bộ rồi mới mời tư vấn giám sát nghiệm thu; chỉ chuyển bước khi biên bản được ký (xem BV-03).",
              "Hồ sơ chất lượng lập theo mẫu của tư vấn: phiếu yêu cầu nghiệm thu, biên bản nghiệm thu vật liệu đầu vào, biên bản lấy mẫu thí nghiệm, biên bản nghiệm thu công việc, nghiệm thu giai đoạn cọc, checklist kiểm tra, bảng tổng hợp thi công cọc.",
              "Danh sách thí nghiệm và hồ sơ kèm theo (cọc, cốt thép, bê tông M300 theo từng đợt đổ R7 - R28, độ chặt K95 theo từng ngày đắp, đá xô) theo bảng theo dõi trong hồ sơ chất lượng.",
              "Nhật ký thi công ghi hằng ngày; nhật ký ép cọc, nhật ký bảo dưỡng bê tông và các biên bản kiểm tra lưu đầy đủ.",
              "Xử lý không phù hợp: dừng thi công hạng mục không đạt, lập biên bản, nêu biện pháp khắc phục và nghiệm thu lại; báo cáo tư vấn giám sát và Chủ đầu tư.",
              "Hoàn công: lập bản vẽ hoàn công, tổng hợp hồ sơ chất lượng trước khi nghiệm thu hoàn thành toàn gói và bàn giao."]:
        B(t)

    # ---------------- 8
    H1("8. BIỆN PHÁP AN TOÀN LAO ĐỘNG")
    for t in ["Huấn luyện an toàn cho công nhân trước khi vào công trường; trang bị mũ, giày, găng, áo phản quang; dây an toàn khi làm việc trên cao hoặc gần mép hố.",
              "Rào chắn, biển báo quanh hố móng, đê quai, lạch; có thang lên xuống hố; lắp đèn chiếu sáng khi làm việc ban đêm.",
              "Máy móc thiết bị được kiểm định/kiểm tra trước khi sử dụng, người vận hành có chứng chỉ; vùng hoạt động của cẩu, máy ép cọc có người cảnh giới.",
              "An toàn điện: tủ điện có chống giật, dây điện treo cao hoặc bọc cách điện, không để nước ngập thiết bị điện.",
              "An toàn khi đào: mái hố ổn định, không chất tải sát mép, theo dõi sạt lở sau mưa; không làm việc dưới hố khi có nguy cơ sập.",
              "An toàn khi ép cọc, cẩu cọc: kiểm tra cáp, móc, đối trọng; không đứng dưới tải; cấm người vào vùng nguy hiểm.",
              "An toàn khi đổ bê tông: kiểm tra ván khuôn, giàn giáo, cây chống; cần bơm có người điều khiển, tín hiệu thống nhất.",
              "Sơ cấp cứu: tủ thuốc, danh sách số điện thoại khẩn cấp (y tế, cứu hỏa, chỉ huy trưởng) niêm yết tại công trường."]:
        B(t)

    # ---------------- 9
    H1("9. VỆ SINH MÔI TRƯỜNG, PHÒNG CHÁY CHỮA CHÁY, AN NINH")
    for t in ["Hố lắng trước khi xả nước hố móng; không xả bùn, nhớt ra nguồn nước; thu gom rác, nhiên liệu, dầu thải đúng nơi quy định.",
              "Rửa xe trước khi ra cổng; che phủ xe chở đất, đá; tưới nước giảm bụi trên đường công vụ; hạn chế tiếng ồn, thi công ban đêm chỉ khi được phép.",
              "Bình chữa cháy tại kho, khu hàn, khu nhiên liệu; cấm hút thuốc gần nơi chứa nhiên liệu; có nội quy PCCC.",
              "Có người trực bảo vệ, quản lý người ra vào công trường; kiểm kê vật tư thiết bị cuối ngày; kê biển cấm người ngoài vào."]:
        B(t)

    # ---------------- 10
    H1("10. BIỆN PHÁP ỨNG PHÓ RỦI RO")
    T(["Rủi ro", "Biện pháp"],
      [["Mưa lớn, nước dâng, ngập hố móng", "Theo dõi dự báo; bơm thoát nước liên tục; che phủ mặt đổ bê tông và mái hố; dừng đổ bê tông khi mưa lớn; dự phòng máy bơm, nhiên liệu"],
       ["Sạt lở mái hố, đê quai", "Kiểm tra sau mưa; gia cố mái; thu hẹp đốt đào; sơ tán người và thiết bị khi có dấu hiệu bất thường"],
       ["Cọc ép không đạt độ chối/lực ép, cọc nghiêng, gãy", "Dừng ép, báo thiết kế và tư vấn giám sát; xử lý theo chỉ định (ép bổ sung, nối, cọc thay thế)"],
       ["Bê tông không đạt cường độ mẫu", "Báo tư vấn và thiết kế; khoan lấy mẫu hoặc thí nghiệm bổ sung; xử lý theo chỉ định"],
       ["Chậm tiến độ phân đoạn găng (PĐ4)", "Tăng ca, bổ sung tổ đội và thiết bị, điều chỉnh thứ tự công việc không găng, báo Chủ đầu tư"],
       ["Mất điện, hỏng máy", "Máy phát điện dự phòng, thiết bị dự phòng và nhà cung cấp bê tông thay thế"]], [5.5, 12.5])

    # ---------------- 11
    H1("11. KẾT LUẬN")
    P("Biện pháp thi công nêu trên được lập trên cơ sở tiến độ và lịch nghiệm thu đã thống nhất, phù hợp trình tự thi công cuốn chiếu bốn phân đoạn và các yêu cầu nghiệm thu của dự án. Trong quá trình thi công, nhà thầu sẽ điều chỉnh, bổ sung biện pháp khi có thay đổi về thiết kế, điều kiện hiện trường hoặc theo yêu cầu của tư vấn giám sát và Chủ đầu tư, và chỉ thực hiện sau khi được chấp thuận.", align="j")

    # ---------------- phụ lục
    H1("PHỤ LỤC: DANH MỤC BẢN VẼ VÀ TÀI LIỆU KÈM THEO")
    T(["Số hiệu", "Tên bản vẽ / tài liệu"],
      [["BV-01", "Sơ đồ mặt bằng tổ chức thi công"], ["BV-02", "Mặt cắt ngang điển hình thi công kè"], ["BV-03", "Sơ đồ trình tự thi công và các điểm nghiệm thu"],
       ["BV-04", "Biện pháp thi công ép cọc BTCT 300x300"], ["BV-05", "Biện pháp thi công móng, ván khuôn và đổ bê tông"],
       ["BV-06", "Biện pháp đắp hoàn trả, đá xô, an toàn và môi trường"], ["Tiến độ", "Tiến độ thi công (PDF A3, Excel, MS Project)"]], [3, 14.5], 12, (0,))
    P("Lưu ý: các bản vẽ là sơ đồ minh họa không theo tỉ lệ; kích thước theo hồ sơ thiết kế được duyệt.", italic=True)
    d.save(out)
    print("OK", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
