"""Xuất checklist kiểm tra công tác (.xlsx) đúng form gốc của file mẫu XD / MEP.

Mỗi biên bản -> một file chỉ chứa sheet mẫu đã chọn. Ô đánh giá L1 được tick theo
kết quả người kiểm tra đã nhập; mục "Chưa kiểm tra" để ô trống. Cột L2 (TVGS) luôn để trống.
"""
import copy
import io
import json
import re
import zipfile
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = Path(__file__).parent / "data" / "qaqc"
TICK, EMPTY = "þ", "o"          # Wingdings: ô đã chọn / ô trống
UTICK, UEMPTY = "☑", "☐"        # dùng cho dòng kết luận (chữ Times New Roman)


def _norm(s):
    return " ".join(str(s or "").split()).lower()


def _put(ws, r, c, v):
    cell = ws.cell(r, c)
    if type(cell).__name__ != "MergedCell":
        cell.value = v


def build_checklist_xlsx(rec, project, so_kt=None, suffix=""):
    """Trả về (tên_file, bytes) cho một biên bản có mau_checklist."""
    cls = json.loads((ROOT / "checklists.json").read_text(encoding="utf-8"))
    info = cls[rec["mau_checklist"]]
    wb = openpyxl.load_workbook(ROOT / "templates" / f"checklist_{info['nhom_file']}.xlsx")
    for n in list(wb.sheetnames):
        if n != info["sheet"]:
            del wb[n]
    ws = wb[info["sheet"]]
    d = date.fromisoformat(rec["ngay"])
    ngay = f"{d.day:02d}/{d.month:02d}/{d.year}"

    # Tiêu đề: giữ nhãn của mẫu, thay phần giá trị.
    for row in ws.iter_rows(min_row=1, max_row=8):
        for c in row:
            v = str(c.value or "").strip()
            if v.startswith("Công trình") and project.get("ten_cong_trinh"):
                c.value = f"Công trình: {project['ten_cong_trinh']}"
            elif v.startswith("Gói thầu") and project.get("goi_thau"):
                c.value = f"Gói thầu: {project['goi_thau']}"
            elif v.startswith("Địa điểm") and project.get("dia_diem"):
                c.value = f"Địa điểm: {project['dia_diem']}"
            elif v.startswith("Số:"):
                c.value = f"Số: {so_kt or rec['so'].split(' ', 1)[-1] + '/KT'}"
            elif v.startswith("Ngày"):
                c.value = f"Ngày: {ngay}"
            elif v.startswith("Đối tượng kiểm tra") and rec.get("hang_muc"):
                c.value = f"Đối tượng kiểm tra: {rec['hang_muc']}"
            elif v.upper().startswith(("TÊN, VỊ TRÍ", "TÊN, VỊ TRÍ CẤU KIỆN")) and rec.get("vi_tri") and c.column > 1:
                c.value = f"{v.rstrip(': ')}: {rec['vi_tri']}"

    # Dòng tiêu đề bảng kiểm tra.
    head = next(r for r in range(1, 40) if str(ws.cell(r, 2).value or "").strip().startswith("Công việc kiểm tra"))
    # Cột "Đạt" của L1: mẫu XD "Bê tông lót", "Ván khuôn", "Đổ bê tông" lệch 1 cột so với các mẫu khác.
    c0 = next((c.column for rr in (head, head + 1) for c in ws[rr] if str(c.value or "").strip() == "Đạt"), 5)
    results = {}
    for it in rec["checklist"]:
        results.setdefault(_norm(it["Nội dung"]), []).append(it)
    ghi_chu = []
    for r in range(head + 2, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        b = str(ws.cell(r, 2).value or "")
        if a.startswith(("Ý kiến", "KẾT LUẬN")):
            ykien_row = r if a.startswith("Ý kiến") else None
            if ykien_row:
                break
            continue
        if re.fullmatch(r"\.{0,}\s*/\s*\.{0,}\s*/.*", str(ws.cell(r, 3).value or "").replace("…", ".")) or "/..." in str(ws.cell(r, 3).value or ""):
            _put(ws, r, 3, ngay)
        if a.isdigit() and b.strip():
            lst = results.get(_norm(b))
            it = lst.pop(0) if lst else None
            for col in range(c0, c0 + 4):
                _put(ws, r, col, EMPTY)
            if it:
                kq = it.get("Kết quả", "")
                if kq == "Đạt":
                    _put(ws, r, c0, TICK)
                elif kq.startswith("Không đạt"):
                    _put(ws, r, c0 + 1, TICK)
                note = str(it.get("Ghi chú") or "").strip()
                if note or kq == "Đạt có điều kiện":
                    ghi_chu.append(f"- {b.strip()}: {kq if kq == 'Đạt có điều kiện' else ''} {note}".rstrip(": ").strip())
    # Ý kiến khác: chỉ ghi những gì người kiểm tra đã nhập.
    for r in range(head, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("Ý kiến") and ghi_chu:
            ws.cell(r, 1).value = str(ws.cell(r, 1).value).rstrip() + "\n" + "\n".join(ghi_chu)
            if not any(r in range(m.min_row, m.max_row + 1) and m.min_col == 1 for m in ws.merged_cells.ranges):
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 16 * (1 + len(ghi_chu))
            break

    # Kết luận: chỉ tick khi người dùng đã chọn kết luận.
    kl = rec.get("ket_luan", "")
    for r in range(head, ws.max_row + 1):
        c = ws.cell(r, 2)
        v = str(c.value or "")
        m = re.match(r"^\s*\S\s+(Đồng ý|Không đồng ý) chuyển công tác tiếp theo", v)
        if m:
            tick = (m.group(1) == "Đồng ý" and kl in ("Đạt", "Đạt có điều kiện")) or \
                   (m.group(1) == "Không đồng ý" and kl.startswith("Không đạt"))
            c.value = f"{UTICK if tick else UEMPTY} {m.group(1)} chuyển công tác tiếp theo"
            f = copy.copy(c.font)
            f.name = "Times New Roman"
            c.font = f

    # Họ tên người ký ở hai dòng cuối: nhà thầu = chỉ huy trưởng, TVGS = cán bộ kỹ thuật (chữ ký để trống)
    for r in range(head, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        if a.startswith("Nhà thầu thi công") and project.get("chi_huy_truong"):
            _put(ws, r, 3, project["chi_huy_truong"])
            ws.row_dimensions[r].height = 45  # rộng chỗ để ký
        elif a.startswith("Tư vấn giám sát") and project.get("can_bo_tvgs"):
            _put(ws, r, 3, project["can_bo_tvgs"])
            ws.row_dimensions[r].height = 45
    logo = project.get("logo")
    if logo and (ROOT / "templates" / logo).exists():  # thay ô "LOGO NHÀ THẦU" (ảnh ở cột G, hàng 2)
        # ô "LOGO NHÀ THẦU" là ảnh ngoài cùng bên phải ở hàng 2 (cột G hoặc H tùy mẫu)
        hang2 = [i for i in ws._images if i.anchor._from.row == 1]
        old = max(hang2, key=lambda i: i.anchor._from.col) if hang2 else None
        if old is not None:
            new = Image(str(ROOT / "templates" / logo))
            new.anchor, new.width, new.height = old.anchor, old.width, old.height
            ws._images.remove(old)
            ws.add_image(new)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, (1 if info["nhom_file"] == "XD" else 0)
    buf = io.BytesIO()
    wb.save(buf)
    return f"Checklist_{rec['so'].replace(' ', '_')}{suffix}.xlsx", buf.getvalue()


def build_checklists_zip(records, project):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for rec in records:
            name, data = build_checklist_xlsx(rec, project)
            z.writestr(name, data)
    return buf.getvalue()
