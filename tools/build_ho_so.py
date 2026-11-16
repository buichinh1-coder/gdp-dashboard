"""Dựng bộ hồ sơ nghiệm thu từ lịch nghiệm thu (xlsx) -> records.json + file Word/Excel.

Ngày nghiệm thu = ngày do người dùng xác nhận là thực tế (lấy theo ngày kế hoạch).
Kết quả kiểm tra, kết luận, tên người ký: để trống, người kiểm tra điền.
Dùng: python tools/build_ho_so.py <lich.xlsx> <thư_mục_ra>
"""
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

raw = []  # (ngay, stt, rec)
vl = [  # ngày VL do người lập chọn: cọc = ngày đầu tiên trước ép; thép = trước ngày bắt đầu thép; BT = ngày bắt đầu đổ
    ("Cọc bê tông BT300x300 M300 bền sunfat (tự sản xuất)", "Toàn tuyến cọc 88-126", "m", 3372, "2026-08-24"),
    ("Cốt thép", "Toàn tuyến cọc 88-126", "tấn", 23.02, "2026-09-04"),
]
bt = {"Cọc 88-98": ("2026-09-11", 136.94), "Cọc 98-108": ("2026-09-16", 136.94), "Cọc 108-118": ("2026-09-21", 136.94), "Cọc 118-126": ("2026-09-26", 136.96)}
for pd_, (d, q) in bt.items():
    vl.append(("Bê tông thương phẩm M300 bền sunfat, phụ gia chống thấm (cấp phối đã duyệt)", pd_, "m3", q, d))
for k, (hm, vt, dv, kl, d) in enumerate(vl):
    raw.append((d, k, {"type": "VL", "hang_muc": hm, "vi_tri": vt, "dvt": dv, "khoi_luong": kl, "ngay": d, "mau_checklist": ""}))
for v in rows:
    muc = str(v[1] or "")
    if muc == "01" or not muc:  # cung cấp cọc -> VL đầu vào 1 lần; hoàn thành/bàn giao -> form khác, chưa hỗ trợ
        continue
    raw.append((iso(v[8]), int(v[0]), {"type": "CV", "hang_muc": v[2], "vi_tri": v[3], "dvt": v[4], "khoi_luong": float(v[5] or 0),
                                         "ngay": iso(v[8]), "mau_checklist": MAU.get(muc, ""), "muc_hd": muc}))
db = {"project": {"ten_cong_trinh": "Kè bê tông phía núi, cọc 88–126", "so_hop_dong": "12.08/2026/HĐKT/ĐGC-GL", "chu_dau_tu": "",
                  "tu_van_giam_sat": "", "nha_thau": "", "dia_diem": "", "goi_thau": ""}, "records": []}
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
# danh mục
TY = {"VL": "Nghiệm thu vật liệu đầu vào", "CV": "Nghiệm thu công việc xây dựng"}
dm = pd.DataFrame([{"Số": r["so"], "Loại": TY[r["type"]], "Ngày nghiệm thu": r["ngay"], "Hạng mục": r["hang_muc"], "Vị trí": r["vi_tri"],
                    "ĐVT": r["dvt"], "Khối lượng": r["khoi_luong"], "Mục HĐ": r.get("muc_hd", ""),
                    "Mẫu checklist": CL[r["mau_checklist"]]["doi_tuong"] if r["mau_checklist"] else "(chưa có mẫu)",
                    "Kết luận": r["ket_luan"]} for r in recs])
dm.to_excel(OUT / "00_Danh_muc_ho_so.xlsx", index=False)
print(cnt, "file:", sum(1 for _ in OUT.rglob("*") if _.is_file()))
