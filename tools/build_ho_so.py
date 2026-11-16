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
import tieu_chuan  # noqa: E402

ROOT = Path(__file__).parent.parent
OUT = Path(sys.argv[2])
CL = json.loads((ROOT / "data/qaqc/checklists.json").read_text(encoding="utf-8"))

# Mục HĐ -> mẫu checklist gần nhất (cần người dùng xác nhận). Không có mẫu -> không xuất checklist.
MAU = {
    "15": "XD|San lẮP", "16": "XD|đào đất", "04": "XD|đào đất", "08": "XD|Bê tông lot",
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

# Thứ tự công nghệ để xếp các biên bản cùng ngày: đê quai, nạo vét, ép cọc, đào móng, vận chuyển, cắt đầu cọc,
# ván khuôn lót, BT lót, thép, ván khuôn, khe lún, đổ BT, đắp hoàn trả, đá xô, biện pháp thi công
TECH = ["15", "16", "02+03", "04", "05", "10", "08", "13", "11", "12", "14", "09", "06", "07", "17"]
raw = []  # (ngay, thứ tự, rec); VL (thứ tự < 100) xếp trước công việc cùng ngày
vl = [  # ngày VL do người lập chọn: cọc = ngày đầu tiên trước ép; thép = trước ngày bắt đầu thép; BT = ngày bắt đầu đổ
    ("Cọc bê tông BT300x300 M300 bền sunfat (tự sản xuất)", "Toàn tuyến cọc 88-126", "m", 3372, "2026-08-24"),
    ("Cốt thép", "Toàn tuyến cọc 88-126", "tấn", 23.02, "2026-09-04"),
]
for k, (hm, vt, dv, kl, d) in enumerate(vl):
    raw.append((d, k, {"type": "VL", "hang_muc": hm, "vi_tri": vt, "dvt": dv, "khoi_luong": kl, "ngay": d, "mau_checklist": "", "lay_mau": True}))
# B: gộp ép cọc (02) và cắt đầu cọc (03) thành một biên bản "nghiệm thu cọc" cho mỗi phân đoạn
for pd_ in ["Cọc 88-98", "Cọc 98-108", "Cọc 108-118", "Cọc 118-126"]:
    ep = next(v for v in rows if str(v[1]) == "02" and v[3] == pd_)
    cat = next(v for v in rows if str(v[1]) == "03" and v[3] == pd_)
    ngay_ = max(iso(ep[8]), iso(cat[8]))
    raw.append((ngay_, TECH.index("02+03") + 100, {
        "type": "CV", "hang_muc": "Ép cọc BT300x300 (kiểm soát lực ép, nhật ký ép); cắt, đập đầu cọc, thu gom và vận chuyển",
        "vi_tri": pd_, "dvt": "", "khoi_luong": 0, "ngay": ngay_, "mau_checklist": "XD|Ép cọc", "mau_checklist_2": "XD|Cắt đầu cọc",
        "muc_hd": "02+03", "kl_ep": float(ep[5]), "so_coc": int(float(cat[5])), "ep_tu": iso(ep[6]), "ep_den": iso(ep[7])}))
for v in rows:
    muc = str(v[1] or "")
    if muc in ("01", "02", "03") or not muc:  # cung cấp cọc -> VL đầu vào 1 lần; hoàn thành/bàn giao -> form khác, chưa hỗ trợ
        continue
    ngay, kh = iso(v[8]), None
    rec_ = {"type": "CV", "hang_muc": v[2], "vi_tri": v[3], "dvt": v[4], "khoi_luong": float(v[5] or 0),
            "ngay": ngay, "mau_checklist": MAU.get(muc, ""), "muc_hd": muc}
    if muc == "09":  # lấy mẫu bê tông theo từng đợt đổ, thí nghiệm R7 và R28 nằm trong nghiệm thu công việc đổ bê tông
        rec_["lay_mau"] = True
        rec_["ngay_lay_mau"] = iso(v[6])  # lấy mẫu = ngày đổ bê tông
        rec_["ngay_lay_mau_den"] = iso(v[7])
    if muc in ("06", "15") and "K95" not in rec_["hang_muc"]:  # K95 chỉ áp dụng công tác đắp đất (hoàn trả hố móng, đê quai)
        rec_["hang_muc"] = rec_["hang_muc"].rstrip() + " (độ chặt K95)"
    if kh:
        rec_["ngay_ke_hoach"] = kh
    raw.append((ngay, TECH.index(muc) + 100, rec_))
db = {"project": {
    "ten_cong_trinh": "Khu đô thị tại khu 10B, phường Quang Hanh, thành phố Cẩm Phả, tỉnh Quảng Ninh",
    "so_hop_dong": "12.08/2026/HĐKT/ĐGC-GL", "chu_dau_tu": "Công ty TNHH Đỗ Gia Capital", "tu_van_giam_sat": "",
    "nha_thau": "Công ty TNHH Giang Long",
    "dia_diem": "Khu đô thị tại khu 10B, phường Quang Hanh, thành phố Cẩm Phả, tỉnh Quảng Ninh",
    "goi_thau": "Thi công kè chắn đất bê tông đổ tại chỗ tuyến 3 đoạn từ cọc 88 đến cọc 126",
    "chi_huy_truong": "Nguyễn Mạnh Tùng", "can_bo_tvgs": "Nguyễn Thế Anh", "giam_sat_truong": "Nguyễn Huy Hà",
    "logo": "logo_giang_long.png"}, "records": []}
cnt = {"VL": 0, "CV": 0}
for ngay, stt, rec in sorted(raw, key=lambda x: (x[0], x[1])):
    cnt[rec["type"]] += 1
    rec["so"] = f"{rec['type']} {cnt[rec['type']]:03d}"
    rec["checklist"] = checklist(rec["type"], rec["mau_checklist"])
    if rec.get("mau_checklist_2"):
        rec["checklist_2"] = checklist(rec["type"], rec["mau_checklist_2"])
    rec["ket_luan"] = "Chưa có kết luận"
    db["records"].append(rec)
# --- số tham chiếu giữa các biểu mẫu ---
vls = sorted([r for r in db["records"] if r["type"] == "VL"], key=lambda r: (r["ngay"], r["so"]))
lm_src = sorted([r for r in db["records"] if r.get("lay_mau")], key=lambda r: (r.get("ngay_lay_mau") or r["ngay"], r["so"]))
for i, r in enumerate(lm_src, 1):
    r["so_lm"] = f"LM {i:03d}"
num = lambda r: r["so"].split()[-1]
VLREF = {"02+03": lambda r: [x for x in vls if "Cọc bê tông" in x["hang_muc"]],
         "13": lambda r: [x for x in vls if "Cốt thép" in x["hang_muc"]],
         }
for r in db["records"]:
    r["tieu_chuan"] = tieu_chuan.tieu_chuan_vl(r["hang_muc"]) if r["type"] == "VL" else tieu_chuan.tieu_chuan_cv(r.get("muc_hd", ""))
    if r["type"] == "VL":
        r["lm_ref"] = [num(r_) + "/LM" for r_ in [r]] and [f"{r['so_lm'].split()[-1]}/LM"]
        r["so_yc"] = f"{num(r)}/YC/VL-TB"
    else:
        r["so_yc"] = f"{num(r)}/YC/CV"
        if r.get("mau_checklist_2"):
            r["so_kt"] = f"{num(r)}/KT-1 (ép cọc), {num(r)}/KT-2 (cắt đầu cọc)"
        elif r["mau_checklist"]:
            r["so_kt"] = f"{num(r)}/KT"
        if r.get("lay_mau"):
            r["lm_ref"] = [f"{r['so_lm'].split()[-1]}/LM"]
        refs = VLREF.get(r["muc_hd"], lambda r: [])(r)
        if refs:
            r["vl_ref"] = [f"{num(x)}/CV/VL-TB" for x in refs]
            r["lm_ref"] = [f"{x['so_lm'].split()[-1]}/LM" for x in refs if x.get("lay_mau")]
coc = [r for r in db["records"] if r["type"] == "CV" and r.get("muc_hd") == "02+03"]
vl_coc = next(x for x in vls if "Cọc bê tông" in x["hang_muc"])
gd = {"type": "GD", "so": "GD 001", "ngay": max(r["ngay"] for r in coc), "vi_tri": "Toàn tuyến cọc 88-126", "dvt": "", "khoi_luong": 0,
      "hang_muc": "Thi công cọc BT300x300 toàn tuyến (4 phân đoạn): ép cọc, cắt, đập đầu cọc",
      "mau_checklist": "", "ket_luan": "Chưa có kết luận", "checklist": [], "so_yc": "001/YC/GĐ",
      "cv_ref": [f"{num(r)}/CV" for r in coc], "vl_ref": [f"{num(vl_coc)}/CV/VL-TB"], "lm_ref": [f"{vl_coc['so_lm'].split()[-1]}/LM"],
      "tieu_chuan": tieu_chuan.tieu_chuan_cv("02+03"), "bang_kl": "Bảng tổng hợp công tác thi công cọc (đính kèm)"}
db["records"].append(gd)
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
        if r.get("mau_checklist_2"):  # nghiệm thu cọc: 2 checklist (ép cọc, cắt đầu cọc)
            n1, b1 = qaqc_xlsx.build_checklist_xlsx(r, proj, so_kt=f"{num(r)}/KT-1", suffix="_ep_coc")
            (sub / n1).write_bytes(b1)
            r2 = dict(r, mau_checklist=r["mau_checklist_2"], checklist=r["checklist_2"])
            n2, b2 = qaqc_xlsx.build_checklist_xlsx(r2, proj, so_kt=f"{num(r)}/KT-2", suffix="_cat_dau_coc")
            (sub / n2).write_bytes(b2)
        else:
            n, b = qaqc_xlsx.build_checklist_xlsx(r, proj)
            (sub / n).write_bytes(b)
# C: nghiệm thu giai đoạn cọc (toàn tuyến) + bảng tổng hợp công tác thi công cọc
gd = next(r for r in recs if r["type"] == "GD")
(OUT / "03_Nghiem_thu_giai_doan_coc.docx").write_bytes(qaqc_docx.build_docx([gd], proj))
wb_th = openpyxl.load_workbook(ROOT / "data/qaqc/templates/tonghop_ep_coc.xlsx")
ws_th = wb_th.worksheets[0]
ws_th["A3"] = f"Công trình: {proj['ten_cong_trinh']}"
ws_th["A4"] = f"Gói thầu: {proj['goi_thau']}"
ws_th["A5"] = "Hạng mục: Cọc BT300x300 M300 bền sunfat, toàn tuyến cọc 88-126"
ws_th["A6"] = f"Nhà thầu thi công: {proj['nha_thau']}"
coc_recs = sorted([r for r in recs if r.get("muc_hd") == "02+03"], key=lambda r: PD.index(r["vi_tri"]))
for i, r in enumerate(coc_recs):
    row = 10 + i
    ws_th.cell(row, 1).value = i + 1
    ws_th.cell(row, 2).value = f"{r['vi_tri']} ({r['so_coc']} cọc)"
    ws_th.cell(row, 3).value = f"{dt.date.fromisoformat(r['ep_tu']):%d/%m} - {dt.date.fromisoformat(r['ep_den']):%d/%m/%Y}"
    ws_th.cell(row, 4).value = r["kl_ep"]  # khối lượng ép theo lịch; chiều dài từng cọc, cao độ, lực ép điền từ nhật ký ép cọc
ws_th.cell(14, 2).value = f"Tổng cộng ({sum(r['so_coc'] for r in coc_recs)} cọc)"
ws_th.cell(14, 4).value = sum(r["kl_ep"] for r in coc_recs)
from openpyxl.styles import Alignment
for row in range(10, 15):
    ws_th.row_dimensions[row].height = 48
    for col in (2, 3):
        ws_th.cell(row, col).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws_th.cell(row, 4).number_format = "#,##0"
wb_th.save(OUT / "03_Bang_tong_hop_thi_cong_coc.xlsx")
# --- phiếu yêu cầu nghiệm thu, biên bản lấy mẫu ---
(OUT / "05_Phieu_yeu_cau_nghiem_thu").mkdir(exist_ok=True)
(OUT / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_vat_lieu.docx").write_bytes(
    qaqc_docx.build_docx(qaqc_docx.expand([r for r in recs if r["type"] == "VL"], ("yc",)), proj))
for i, p in enumerate(PD, 1):
    sel = qaqc_docx.expand([r for r in recs if r["type"] == "CV" and r["vi_tri"] == p], ("yc",))
    (OUT / f"05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_cong_viec_PD{i}_{p.replace(' ', '_')}.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
sel = qaqc_docx.expand([r for r in recs if r["type"] == "CV" and r["vi_tri"] not in PD], ("yc",))
(OUT / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_cong_viec_toan_tuyen.docx").write_bytes(qaqc_docx.build_docx(sel, proj))
(OUT / "05_Phieu_yeu_cau_nghiem_thu/YC_nghiem_thu_giai_doan_coc.docx").write_bytes(qaqc_docx.build_docx(qaqc_docx.expand([gd], ("yc",)), proj))
(OUT / "06_Bien_ban_lay_mau.docx").write_bytes(qaqc_docx.build_docx(qaqc_docx.expand([r for r in recs if r.get("lay_mau")], ("lm",)), proj))

# --- danh sách kết quả thí nghiệm / hồ sơ cần có (bảng theo dõi) ---
D = lambda x: x.date()
fmt = lambda d: d.strftime("%d/%m/%Y")
rng = lambda a_, b_: fmt(a_) if a_ == b_ else f"{fmt(a_)} - {fmt(b_)}"
row_of = lambda muc, pd_: next(v for v in rows if str(v[1]) == muc and v[3] == pd_)
so_vl = lambda k: next(r["so"] for r in recs if r["type"] == "VL" and k in r["hang_muc"])
cv_of = lambda muc, pd_: next(r for r in recs if r["type"] == "CV" and r.get("muc_hd") == muc and r["vi_tri"] == pd_)
TC = "Theo chỉ dẫn kỹ thuật của hợp đồng"
T = []  # (nhóm, đối tượng, nội dung, tiêu chuẩn, biên bản liên quan, ngày thực hiện / lấy mẫu, ngày cần có kết quả)


def add(nhom, dt_, nd, tc, bb, ngay, kq):
    T.append((nhom, dt_, nd, tc, bb, ngay, kq))


GA, GB = "A. Cọc BT300x300", "B. Cốt thép"
for nd, tc in [("Phiếu duyệt cấp phối / trialmix bê tông cọc", TC), ("Chứng chỉ và kết quả thí nghiệm xi măng (bền sunfat), phụ gia", TC),
               ("Thí nghiệm cát, đá, nước trộn bê tông", "TCVN 7570:2006, TCVN 4506:2012"),
               ("Thí nghiệm kéo, uốn thép chế tạo cọc (nếu cọc có cốt thép)", "TCVN 1651-2:2018, QCVN 07:2019/BKHCN"),
               ("Nén mẫu bê tông cọc R7 (căn cứ cho phép ép; đủ cho toàn bộ cọc)", "TCVN 3118:2022"),
               ("Biên bản kiểm tra chế tạo cọc: kích thước, độ thẳng, mũi cọc, vết nứt", TC)]:
    add(GA, "Cọc toàn tuyến", nd, tc, f"{so_vl('Cọc')} / LM 001", "Trước ép cọc", "Trước 24/08/2026")
for nd, tc in [("CO/CQ và chứng chỉ xuất xưởng từng lô thép", "TCVN 1651-2:2018"),
               ("Thí nghiệm kéo (giới hạn chảy, bền, độ giãn dài) và uốn nguội theo lô, đường kính", "TCVN 1651-2:2018, QCVN 07:2019/BKHCN")]:
    add(GB, "Thép toàn tuyến", nd, tc, f"{so_vl('Cốt thép')} / LM 002", "Trước lắp dựng", "Trước 04/09/2026")
GC = "C. Bê tông M300 (từng đợt đổ)"
first_pour = min(D(row_of("09", p_)[6]) for p_ in PD)
add(GC, "Toàn bộ các đợt", "Hồ sơ cấp phối đã duyệt, chứng chỉ vật liệu của nhà máy bê tông", TC, "-", "Trước đợt đổ đầu", f"Trước {fmt(first_pour)}")
for p_ in PD:
    r = row_of("09", p_)
    b0, b1 = D(r[6]), D(r[7])
    c = cv_of("09", p_)
    ref = f"{c['so']} / {c['so_lm']}"
    dd = f"Đợt đổ {p_.lower()}"
    add(GC, dd, "Phiếu giao hàng từng xe (mác, độ sụt, giờ trộn)", "-", ref, rng(b0, b1), rng(b0, b1))
    add(GC, dd, "Độ sụt hiện trường", "TCVN 3106:2022", ref, rng(b0, b1), rng(b0, b1))
    add(GC, dd, "Lấy mẫu bê tông tại hiện trường (biên bản lấy mẫu)", "TCVN 3105:2022", ref, rng(b0, b1), rng(b0, b1))
    add(GC, dd, "Nén mẫu R7 (căn cứ sớm để đắp đất lên tường chắn)", "TCVN 3118:2022", ref, rng(b0, b1), rng(b0 + dt.timedelta(7), b1 + dt.timedelta(7)))
    add(GC, dd, "Nén mẫu R28 (căn cứ nghiệm thu kết cấu)", "TCVN 3118:2022", ref, rng(b0, b1), rng(b0 + dt.timedelta(28), b1 + dt.timedelta(28)))
    add(GC, dd, "Thấm nước (bê tông chống thấm), nếu chỉ dẫn kỹ thuật yêu cầu", "TCVN 3116:2022", ref, rng(b0, b1), rng(b0 + dt.timedelta(28), b1 + dt.timedelta(28)))
GD_ = "D. Đắp đất - độ chặt K95"
add(GD_, "Đất đắp đê quai", "Thí nghiệm đầm nén tiêu chuẩn: dung trọng khô lớn nhất, độ ẩm tốt nhất", TC, cv_of("15", PD[0])["so"], "Trước khi đắp", fmt(D(row_of("15", PD[0])[6])))
add(GD_, "Đất đắp hoàn trả hố móng", "Thí nghiệm đầm nén tiêu chuẩn: dung trọng khô lớn nhất, độ ẩm tốt nhất", TC, cv_of("06", PD[0])["so"], "Trước khi đắp", fmt(D(row_of("06", PD[0])[6])))
for muc, ten in (("15", "Đắp đê quai"), ("06", "Đắp hoàn trả hố móng")):
    for p_ in PD:
        r = row_of(muc, p_)
        add(GD_, f"{ten} - {p_.lower()}", "Độ chặt K95 từng lớp đắp (thí nghiệm hiện trường): mỗi lớp đạt K >= 0,95 mới đắp lớp tiếp theo", TC,
            cv_of(muc, p_)["so"], rng(D(r[6]), D(r[7])), "Từng lớp đắp")
GE = "E. Đề xuất (xác nhận)"
add(GE, "Khe lún", "Ống thoát nước, chi tiết thoát nước khe lún: chứng chỉ chất lượng", TC, cv_of("14", PD[0])["so"], "Trước lắp", fmt(D(row_of("14", PD[0])[6])))
add(GE, "Khe lún", "Vật tư khe lún: chứng chỉ, mẫu vật liệu", TC, cv_of("14", PD[0])["so"], "Trước lắp", fmt(D(row_of("14", PD[0])[6])))
add(GE, "Đá xô chân tường chắn", "Thí nghiệm đá", TC, cv_of("07", PD[0])["so"], "Trước khi đắp đá xô", fmt(D(row_of("07", PD[0])[6])))
GF = "F. Ép cọc"
add(GF, "Thiết bị ép cọc", "Giấy kiểm định thiết bị ép (đồng hồ áp lực, kích) còn hiệu lực", "-", "-", "Trước ép cọc", "Trước 25/08/2026")
for p_ in PD:
    c = cv_of("02+03", p_)
    add(GF, p_, "Nhật ký ép từng cọc và bảng tổng hợp công tác ép cọc", "TCVN 9394:2012", c["so"], rng(dt.date.fromisoformat(c["ep_tu"]), dt.date.fromisoformat(c["ep_den"])), c["ngay"][8:] + "/" + c["ngay"][5:7] + "/" + c["ngay"][:4])
add(GF, "Cọc toàn tuyến", "Thí nghiệm nén tĩnh cọc hoặc PDA (nếu thiết kế yêu cầu)", "TCVN 9393:2012", "GD 001", "Theo thiết kế", "Theo thiết kế")
GG = "G. Trắc đạc, hoàn công"
for p_ in PD:
    c = cv_of("02+03", p_)
    add(GG, p_, "Biên bản trắc đạc vị trí cọc, cao độ đầu cọc sau cắt", "-", c["so"], c["ngay"][8:] + "/" + c["ngay"][5:7] + "/" + c["ngay"][:4], "Cùng nghiệm thu cọc")
add(GG, "Toàn tuyến", "Bản vẽ hoàn công, bảng khối lượng hoàn thành", "-", "Nghiệm thu hoàn thành", "Trước hoàn thành", "Trước 27/10/2026")
tn_df = pd.DataFrame([{"STT": i, "Nhóm": a, "Đối tượng": b, "Kết quả thí nghiệm / hồ sơ cần có": c, "Tiêu chuẩn (tham khảo)": d, "Biên bản liên quan": e,
                       "Ngày thực hiện / lấy mẫu": f, "Ngày cần có kết quả": g, "Phòng thí nghiệm": "", "Số phiếu kết quả": "", "Ngày phiếu": "",
                       "Kết quả": "", "Đạt / Không đạt": "", "Ghi chú": ""} for i, (a, b, c, d, e, f, g) in enumerate(T, 1)])
r7 = []
nt_bt = {r["vi_tri"]: r["ngay"] for r in recs if r["type"] == "CV" and r.get("muc_hd") == "09"}
for pd_ in PD:
    r = row_of("09", pd_)
    b0, b1 = D(r[6]), D(r[7])
    dap = D(row_of("06", pd_)[6])
    nt = dt.date.fromisoformat(nt_bt[pd_])
    c9 = cv_of("09", pd_)
    r7.append({"Phân đoạn": pd_, "Biên bản (CV đổ BT / LM)": f"{c9['so']} / {c9['so_lm']}", "Ngày đổ (lấy mẫu)": rng(b0, b1),
               "Ngày nén R7": rng(b0 + dt.timedelta(7), b1 + dt.timedelta(7)), "Ngày nén R28": rng(b0 + dt.timedelta(28), b1 + dt.timedelta(28)),
               "Nghiệm thu kết cấu BT (theo lịch)": fmt(nt), "Ngày bắt đầu đắp hoàn trả": fmt(dap),
               "Nghiệm thu kết cấu sau R28?": "Có" if nt >= b1 + dt.timedelta(28) else "KHÔNG - sớm hơn R28",
               "Đắp trước khi có R28?": "Có - chỉ được đắp khi kết quả R7 đạt và TVGS chấp thuận" if dap < b1 + dt.timedelta(28) else "Không",
               "Ghi chú": "Mỗi đợt đổ lấy mẫu riêng, nén R7 và R28"})
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties
with pd.ExcelWriter(OUT / "07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx", engine="openpyxl") as w:
    tn_df.to_excel(w, sheet_name="Danh_sach", index=False, startrow=3)
    pd.DataFrame(r7).to_excel(w, sheet_name="Moc_R7_R28", index=False)
    pd.DataFrame([{"Số hiệu": a, "Tên": b, "Tình trạng tra cứu": c, "Nguồn": d} for a, b, c, d in tieu_chuan.TC_BANG]).to_excel(w, sheet_name="Tieu_chuan_da_dung", index=False)
    pd.DataFrame([{"Số hiệu / văn bản": a, "Tên": b, "Cần xác nhận": c, "Nguồn": d} for a, b, c, d in tieu_chuan.TC_CAN_XAC_NHAN]).to_excel(w, sheet_name="Tieu_chuan_can_xac_nhan", index=False)
    thin = Side(style="thin", color="9AA5B1")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws_ in w.book.worksheets:
        main = ws_.title == "Danh_sach"
        hr = 4 if main else 1
        for c in ws_[hr]:
            c.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="123E70")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = brd
        ws_.row_dimensions[hr].height = 34
        for row in ws_.iter_rows(min_row=hr + 1):
            for c in row:
                c.font = Font(name="Times New Roman", size=10.5)
                c.alignment = Alignment(wrap_text=True, vertical="center")
                c.border = brd
        ws_.freeze_panes = ws_.cell(hr + 1, 1)
        ws_.page_setup.orientation = "landscape"
        ws_.page_setup.paperSize = ws_.PAPERSIZE_A4
        ws_.page_setup.fitToWidth, ws_.page_setup.fitToHeight = 1, 0
        ws_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws_.page_margins.left = ws_.page_margins.right = 0.4
        ws_.print_title_rows = f"{hr}:{hr}"
    ws = w.book["Danh_sach"]
    ws.merge_cells("A1:N1")
    ws["A1"] = "DANH SÁCH KẾT QUẢ THÍ NGHIỆM, HỒ SƠ CẦN CÓ"
    ws["A1"].font = Font(name="Times New Roman", size=16, bold=True, color="123E70")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:N2")
    ws["A2"] = f"{proj['goi_thau']} - Hợp đồng số {proj['so_hop_dong']} - {proj['nha_thau']}"
    ws["A2"].font = Font(name="Times New Roman", size=11, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A3:N3")
    ws["A3"] = "Các cột Phòng thí nghiệm đến Ghi chú để trống, điền khi có kết quả thực tế. Tiêu chuẩn chỉ để tham khảo, đối chiếu chỉ dẫn kỹ thuật."
    ws["A3"].font = Font(name="Times New Roman", size=10, italic=True)
    for col, wd in zip("ABCDEFGHIJKLMN", (5.5, 12, 20, 40, 20, 14, 18, 18, 12, 10, 10, 9, 8, 12)):
        ws.column_dimensions[col].width = wd
    for name, widths in (("Moc_R7_R28", (14, 20, 22, 22, 22, 22, 20, 20, 42, 30)), ("Tieu_chuan_da_dung", (22, 55, 50, 60)),
                         ("Tieu_chuan_can_xac_nhan", (30, 30, 70, 50))):
        for col, wd in zip("ABCDEFGHIJ", widths):
            w.book[name].column_dimensions[col].width = wd

# danh mục
TY = {"VL": "Nghiệm thu vật liệu đầu vào", "CV": "Nghiệm thu công việc xây dựng", "GD": "Nghiệm thu giai đoạn"}
dm = pd.DataFrame([{"Số": r["so"], "Loại": TY[r["type"]], "Ngày nghiệm thu": r["ngay"], "Hạng mục": r["hang_muc"], "Vị trí": r["vi_tri"],
                    "ĐVT": r["dvt"], "Khối lượng": r["khoi_luong"], "Mục HĐ": r.get("muc_hd", ""), "Ngày kế hoạch ban đầu": r.get("ngay_ke_hoach", ""),
                    "Mẫu checklist": ("; ".join(CL[m]["doi_tuong"] for m in (r["mau_checklist"], r.get("mau_checklist_2")) if m) if r["mau_checklist"] else ("-" if r["type"] == "GD" else "(chưa có mẫu)")),
                    "Kết luận": r["ket_luan"]} for r in recs])
dm.to_excel(OUT / "00_Danh_muc_ho_so.xlsx", index=False)
import shutil; shutil.copy(ROOT / "tools/ho_so_ghi_chu.md", OUT / "01_Ghi_chu_va_gia_dinh.md")
print(cnt, "file:", sum(1 for _ in OUT.rglob("*") if _.is_file()))
