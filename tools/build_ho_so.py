"""Dựng bộ hồ sơ nghiệm thu từ lịch nghiệm thu (xlsx) -> records.json + file Word/Excel.

Ngày nghiệm thu = ngày do người dùng xác nhận là thực tế (lấy theo ngày kế hoạch).
Kết quả kiểm tra, kết luận, tên người ký: để trống, người kiểm tra điền.
Dùng: python tools/build_ho_so.py <lich.xlsx> <thư_mục_ra>
"""
import datetime as dt
import json
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))
import qaqc_docx  # noqa: E402
import qaqc_xlsx  # noqa: E402

ROOT = Path(__file__).parent.parent
OUT = Path(sys.argv[2])
CL = json.loads((ROOT / "data/qaqc/checklists.json").read_text(encoding="utf-8"))

# Mục HĐ -> mẫu checklist gần nhất (cần người dùng xác nhận). Không có mẫu -> không xuất checklist.
MAU = {
    "02": "XD|Ép cọc", "03": "XD|Cắt đầu cọc", "15": "XD|San lẮP", "16": "XD|đào đất", "04": "XD|đào đất", "08": "XD|Bê tông lot",
    "10": "XD|Ván khuôn dầm sàn", "11": "XD|Ván khuôn dầm sàn", "12": "XD|Ván khuôn dầm sàn",
    "13": "XD|Thép dầm sàn ", "09": "XD|Đổ bê tông ", "06": "XD|San lẮP", "07": "XD|San lẮP",
    "14": "XD|khe lún",
}
GENERIC = {"VL": ["Có chứng chỉ xuất xưởng / CO-CQ", "Có kết quả thí nghiệm đầu vào", "Đúng chủng loại, quy cách theo thiết kế"],
           "CV": ["Có bản vẽ thi công được duyệt", "Vật liệu đã được nghiệm thu", "Đúng vị trí, kích thước, cao độ",
                  "Biện pháp thi công / an toàn đã duyệt"]}


def checklist(t, mau):
    if mau:
        return [{"Nhóm": m["nhom"], "Nội dung": m["noi_dung"], "Phương pháp": m["phuong_phap"], "Kết quả": "Chưa kiểm tra", "Ghi chú": ""}
                for m in CL[mau]["muc"]]
    return [{"Nhóm": "", "Nội dung": x, "Phương pháp": "", "Kết quả": "Chưa kiểm tra", "Ghi chú": ""} for x in GENERIC[t]]


ws = openpyxl.load_workbook(sys.argv[1], data_only=True)["Chi tiet nghiem thu"]
rows = [[ws.cell(r, c).value for c in range(1, 15)] for r in range(1, ws.max_row + 1)]
rows = [v for v in rows if str(v[0]).isdigit()]
iso = lambda x: x.date().isoformat()

raw = []  # (ngay, stt, rec)
vl = [  # ngày VL do người lập chọn: cọc = ngày đầu tiên trước ép; thép = trước ngày bắt đầu thép; BT = ngày bắt đầu đổ
    ("Cọc bê tông BT300x300 M300 bền sunfat (tự sản xuất)", "Toàn tuyến cọc 88-126", "m", 3372, "2026-08-24"),
    ("Cốt thép", "Toàn tuyến cọc 88-126", "tấn", 23.02, "2026-09-04"),
]
bt = {"Cọc 88-98": ("2026-09-11", 136.94), "Cọc 98-108": ("2026-09-16", 136.94), "Cọc 108-118": ("2026-09-21", 136.94), "Cọc 118-126": ("2026-09-26", 136.96)}
for pd_, (d, q) in bt.items():
    vl.append(("Bê tông thương phẩm M300 bền sunfat, phụ gia chống thấm (cấp phối đã duyệt)", pd_, "m3", q, d))
for k, (hm, vt, dv, kl, d) in enumerate(vl):
    raw.append((d, k, {"type": "VL", "hang_muc": hm, "vi_tri": vt, "dvt": dv, "khoi_luong": kl, "ngay": d, "mau_checklist": "", "lay_mau": True}))
for v in rows:
    muc = str(v[1] or "")
    if muc == "01" or not muc:  # cung cấp cọc -> VL đầu vào 1 lần; hoàn thành/bàn giao -> form khác, chưa hỗ trợ
        continue
    raw.append((iso(v[8]), int(v[0]), {"type": "CV", "hang_muc": v[2], "vi_tri": v[3], "dvt": v[4], "khoi_luong": float(v[5] or 0),
                                         "ngay": iso(v[8]), "mau_checklist": MAU.get(muc, ""), "muc_hd": muc}))
db = {"project": {
    "ten_cong_trinh": "Khu đô thị tại khu 10B, phường Quang Hanh, thành phố Cẩm Phả, tỉnh Quảng Ninh",
    "so_hop_dong": "12.08/2026/HĐKT/ĐGC-GL", "chu_dau_tu": "Công ty TNHH Đỗ Gia Capital", "tu_van_giam_sat": "",
    "nha_thau": "Công ty TNHH Giang Long",
    "dia_diem": "Khu đô thị tại khu 10B, phường Quang Hanh, thành phố Cẩm Phả, tỉnh Quảng Ninh",
    "goi_thau": "Thi công kè chắn đất bê tông đổ tại chỗ tuyến 3 đoạn từ cọc 88 đến cọc 126",
    "logo": "logo_giang_long.png"}, "records": []}
cnt = {"VL": 0, "CV": 0}
for ngay, stt, rec in sorted(raw, key=lambda x: (x[0], x[1])):
    cnt[rec["type"]] += 1
    rec["so"] = f"{rec['type']} {cnt[rec['type']]:03d}"
    rec["checklist"] = checklist(rec["type"], rec["mau_checklist"])
    rec["ket_luan"] = "Chưa có kết luận"
    db["records"].append(rec)
(ROOT / "data/qaqc/records.json").write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")

# --- xuất file ---
OUT.mkdir(parents=True, exist_ok=True)
proj = db["project"]
recs = db["records"]
PD = ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]
(OUT / "02_Nghiem_thu_vat_lieu_dau_vao.docx").write_bytes(qaqc_docx.build_docx([r for r in recs if r["type"] == "VL"], proj))
for i, p in enumerate(PD, 1):
    sel = [r for r in recs if r["type"] == "CV" and r["vi_tri"] == p]
    (OUT / f"03_Nghiem_thu_cong_viec_PD{i}_{p.replace(' ', '_')}.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
sel = [r for r in recs if r["type"] == "CV" and r["vi_tri"] not in PD]
(OUT / "03_Nghiem_thu_cong_viec_toan_tuyen.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
for r in recs:
    if r["type"] == "CV" and r["mau_checklist"]:
        sub = OUT / "04_Checklist" / (r["vi_tri"].replace(" ", "_") if r["vi_tri"] in PD else "Toan_tuyen")
        sub.mkdir(parents=True, exist_ok=True)
        n, b = qaqc_xlsx.build_checklist_xlsx(r, proj)
        (sub / n).write_bytes(b)
# --- phiếu yêu cầu nghiệm thu, biên bản lấy mẫu ---
(OUT / "05_Phieu_yeu_cau_nghiem_thu").mkdir(exist_ok=True)
(OUT / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_vat_lieu.docx").write_bytes(
    qaqc_docx.build_docx(qaqc_docx.expand([r for r in recs if r["type"] == "VL"], ("yc",)), proj))
for i, p in enumerate(PD, 1):
    sel = qaqc_docx.expand([r for r in recs if r["type"] == "CV" and r["vi_tri"] == p], ("yc",))
    (OUT / f"05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_cong_viec_PD{i}_{p.replace(' ', '_')}.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
sel = qaqc_docx.expand([r for r in recs if r["type"] == "CV" and r["vi_tri"] not in PD], ("yc",))
(OUT / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_cong_viec_toan_tuyen.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
(OUT / "06_Bien_ban_lay_mau.docx").write_bytes(qaqc_docx.build_docx(qaqc_docx.expand([r for r in recs if r["type"] == "VL"], ("lm",)), proj))

# --- danh sách kết quả thí nghiệm / hồ sơ cần có ---
D = lambda x: x.date()
bt_rows = {v[3]: (D(v[6]), D(v[7])) for v in rows if str(v[1]) == "09"}
dap_rows = {v[3]: D(v[6]) for v in rows if str(v[1]) == "06"}
vl_so = {r["hang_muc"][:12] + r["vi_tri"]: r["so"] for r in recs if r["type"] == "VL"}
so_vl = lambda k: next(r["so"] for r in recs if r["type"] == "VL" and k in r["hang_muc"])
so_bt = {r["vi_tri"]: r["so"] for r in recs if r["type"] == "VL" and "Bê tông thương phẩm" in r["hang_muc"]}
LM_SO = {r["so"]: f"LM {i:03d}" for i, r in enumerate([r for r in recs if r["type"] == "VL"], 1)}
TC = "Theo chỉ dẫn kỹ thuật của hợp đồng"
tn = [
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Phiếu duyệt cấp phối / trialmix bê tông cọc", TC, "Trước khi sản xuất cọc", so_vl("Cọc")),
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Chứng chỉ và kết quả thí nghiệm xi măng (bền sunfat), phụ gia", TC, "Trước 24/08/2026", so_vl("Cọc")),
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Thí nghiệm cát, đá, nước trộn bê tông", "TCVN 7570:2006, TCVN 7572:2006, TCVN 4506:2012", "Trước 24/08/2026", so_vl("Cọc")),
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Thí nghiệm kéo, uốn thép chế tạo cọc (nếu cọc có cốt thép)", "TCVN 1651-2:2018, TCVN 197-1:2014, TCVN 198:2008", "Trước 24/08/2026", so_vl("Cọc")),
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Nén mẫu bê tông cọc R7 (căn cứ cho phép ép; phụ gia R7 đạt 100%)", "TCVN 3118:2022", "Trước 24/08/2026 (nghiệm thu VL 1 lần), đủ cho toàn bộ 3.372 m", so_vl("Cọc")),
    ("A. Cọc BT300x300 M300 bền sunfat (tự sản xuất)", "Biên bản kiểm tra chế tạo cọc: kích thước, độ thẳng, mũi cọc, vết nứt", TC, "Trước 24/08/2026", so_vl("Cọc")),
    ("B. Cốt thép", "CO/CQ và chứng chỉ xuất xưởng từng lô thép", "TCVN 1651-2:2018", "Trước 04/09/2026", so_vl("Cốt thép")),
    ("B. Cốt thép", "Thí nghiệm kéo (giới hạn chảy, bền, độ giãn dài) và uốn nguội theo lô, đường kính", "TCVN 1651-2:2018, TCVN 197-1:2014, TCVN 198:2008", "Trước 04/09/2026 (nghiệm thu VL 1 lần)", so_vl("Cốt thép")),
    ("C. Bê tông thương phẩm M300 (mỗi đợt đổ)", "Hồ sơ cấp phối đã duyệt, chứng chỉ vật liệu của nhà máy bê tông", TC, "Trước ngày đổ đợt đầu (11/09/2026)", ", ".join(so_bt.values())),
    ("C. Bê tông thương phẩm M300 (mỗi đợt đổ)", "Phiếu giao hàng từng xe (mác, độ sụt, giờ trộn)", "", "Mỗi đợt đổ", ", ".join(so_bt.values())),
    ("C. Bê tông thương phẩm M300 (mỗi đợt đổ)", "Đo độ sụt hiện trường", "TCVN 3106:2022", "Mỗi đợt đổ", ", ".join(so_bt.values())),
    ("C. Bê tông thương phẩm M300 (mỗi đợt đổ)", "Lấy mẫu và nén mẫu R7 (căn cứ đắp đất lên tường chắn)", "TCVN 3105:2022, TCVN 3118:2022", "R7 sau ngày lấy mẫu (xem sheet Moc_R7)", ", ".join(so_bt.values())),
    ("C. Bê tông thương phẩm M300 (mỗi đợt đổ)", "Nén mẫu R28 / thấm nước (nếu chỉ dẫn kỹ thuật hoặc CĐT yêu cầu)", "TCVN 3118:2022, TCVN 3116:2022", "Theo yêu cầu", ", ".join(so_bt.values())),
    ("D. Đề xuất, cần xác nhận", "Bê tông lót M100: kết quả nén mẫu (nếu trộn hoặc mua riêng)", "TCVN 3118:2022", "Mỗi đợt đổ lót", "CV bê tông lót"),
    ("D. Đề xuất, cần xác nhận", "Đất đắp hoàn trả (8.642 m3): thí nghiệm đầm nén tiêu chuẩn (dung trọng khô lớn nhất, độ ẩm tốt nhất)", TC, "Trước khi đắp phân đoạn 1 (21/09/2026)", "CV đắp hoàn trả"),
    ("D. Đề xuất, cần xác nhận", "Độ chặt K từng lớp đắp (hiện trường)", TC, "Từng lớp đắp", "CV đắp hoàn trả"),
    ("D. Đề xuất, cần xác nhận", "Đá xô chân tường chắn: thí nghiệm đá", TC, "Trước khi đắp đá xô (26/09/2026)", "CV đắp đá xô"),
    ("D. Đề xuất, cần xác nhận", "Vật tư khe lún: chứng chỉ, mẫu vật liệu", TC, "Trước 08/09/2026", "CV khe lún"),
    ("E. Ép cọc", "Giấy kiểm định thiết bị ép (đồng hồ áp lực, kích) còn hiệu lực", "", "Trước 25/08/2026", "CV ép cọc"),
    ("E. Ép cọc", "Nhật ký ép từng cọc và bảng tổng hợp công tác ép cọc", "TCVN 9394:2012", "Sau mỗi phân đoạn ép", "CV ép cọc"),
    ("E. Ép cọc", "Thí nghiệm nén tĩnh cọc hoặc PDA (nếu thiết kế yêu cầu)", "TCVN 9393:2012", "Theo thiết kế", "CV ép cọc"),
    ("F. Trắc đạc, hoàn công", "Biên bản trắc đạc cọc, cao độ đầu cọc sau cắt, bản vẽ hoàn công, bảng khối lượng", "", "Trước nghiệm thu hoàn thành (27/10/2026)", "CV cắt đầu cọc, nghiệm thu hoàn thành"),
]
tn_df = pd.DataFrame([{"STT": i, "Nhóm": a, "Nội dung kết quả / hồ sơ cần có": b, "Tiêu chuẩn (tham khảo)": c, "Thời điểm cần có": d,
                       "Liên quan biên bản": e, "Trạng thái": "Cần bổ sung", "Số hiệu / ngày kết quả": "", "Ghi chú": ""}
                      for i, (a, b, c, d, e) in enumerate(tn, 1)])
r7 = []
for pd_ in PD:
    (b0, b1), dap = bt_rows[pd_], dap_rows[pd_]
    r7.append({"Phân đoạn": pd_, "Biên bản VL": so_bt[pd_], "Ngày đổ": f"{b0:%d/%m/%Y} - {b1:%d/%m/%Y}",
               "Ngày nén R7 (sớm nhất - muộn nhất)": f"{b0 + dt.timedelta(7):%d/%m/%Y} - {b1 + dt.timedelta(7):%d/%m/%Y}",
               "Ngày bắt đầu đắp hoàn trả": f"{dap:%d/%m/%Y}",
               "Đắp sau khi đủ R7?": "Có" if dap >= b1 + dt.timedelta(7) else "KHÔNG - cần xử lý",
               "Ghi chú": "Cần kết quả nén mẫu R7 đạt yêu cầu trước khi đắp"})
with pd.ExcelWriter(OUT / "07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx", engine="openpyxl") as w:
    tn_df.to_excel(w, sheet_name="Danh_sach", index=False)
    pd.DataFrame(r7).to_excel(w, sheet_name="Moc_R7", index=False)
    for ws_ in w.book.worksheets:
        for col in ws_.columns:
            ws_.column_dimensions[col[0].column_letter].width = min(60, max(12, max(len(str(c.value or "")) for c in col) * 0.9))

# danh mục
TY = {"VL": "Nghiệm thu vật liệu đầu vào", "CV": "Nghiệm thu công việc xây dựng"}
dm = pd.DataFrame([{"Số": r["so"], "Loại": TY[r["type"]], "Ngày nghiệm thu": r["ngay"], "Hạng mục": r["hang_muc"], "Vị trí": r["vi_tri"],
                    "ĐVT": r["dvt"], "Khối lượng": r["khoi_luong"], "Mục HĐ": r.get("muc_hd", ""),
                    "Mẫu checklist": CL[r["mau_checklist"]]["doi_tuong"] if r["mau_checklist"] else "(chưa có mẫu)",
                    "Kết luận": r["ket_luan"]} for r in recs])
dm.to_excel(OUT / "00_Danh_muc_ho_so.xlsx", index=False)
import shutil; shutil.copy(ROOT / "tools/ho_so_ghi_chu.md", OUT / "01_Ghi_chu_va_gia_dinh.md")
print(cnt, "file:", sum(1 for _ in OUT.rglob("*") if _.is_file()))
