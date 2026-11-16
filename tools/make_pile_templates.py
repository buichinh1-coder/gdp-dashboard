"""Thêm 2 mẫu checklist TỰ SOẠN (ép cọc, cắt đầu cọc) vào templates/checklist_XD.xlsx.

File mẫu gốc không có checklist cho 2 công việc này, nên nhân bản bố cục form "Thép cột"
(cùng khung, logo, cột đánh giá L1/L2) và thay nội dung kiểm tra. Nội dung chỉ liệt kê hạng mục
cần kiểm tra, không nêu dung sai bằng số; đối chiếu với chỉ dẫn kỹ thuật / TCVN 9394:2012 và yêu cầu của TVGS.
"""
import copy
import io
import json
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image

ROOT = Path(__file__).parent.parent / "data" / "qaqc"
XLSX = ROOT / "templates" / "checklist_XD.xlsx"
BM = "Bằng mắt thường"

NEW = {
    "Ép cọc": ("BIÊN BẢN KIỂM TRA CÔNG TÁC\nÉP CỌC BÊ TÔNG CỐT THÉP", [
        ("Bản vẽ thi công, vị trí cọc được duyệt", BM), ("Biện pháp thi công ép cọc được duyệt", BM),
        ("Dây ngăn, bảng hiệu, lối lên xuống", BM), ("An toàn thiết bị ép, cẩu, nền kê máy", BM),
        ("Mặt bằng, sàn thi công ép cọc", BM), ("Thiết bị ép: giấy kiểm định, đồng hồ áp lực, lực ép tối đa", "Kiểm tra hồ sơ + bằng mắt"),
        ("Cọc đã nghiệm thu đầu vào (chủng loại, ký hiệu, kích thước, ngoại quan)", "Bằng mắt thường + thước thép"),
        ("Vị trí tim cọc trước khi ép", "Bằng máy toàn đạc"), ("Độ thẳng đứng của cọc khi ép", "Dây dọi / máy kinh vĩ"),
        ("Lực ép, chiều dài ép, điều kiện dừng ép theo thiết kế", "Đối chiếu nhật ký ép cọc"),
        ("Nhật ký ép cọc ghi đầy đủ, có xác nhận", "Kiểm tra hồ sơ"),
        ("Gởi trục, định vị tim cọc", "Bằng máy toàn đạc"), ("Vị trí cọc sau khi ép", "Bằng máy toàn đạc"),
        ("Cao độ đầu cọc sau khi ép", "Bằng máy thủy bình")]),
    "Cắt đầu cọc": ("BIÊN BẢN KIỂM TRA CÔNG TÁC\nCẮT, ĐẬP ĐẦU CỌC", [
        ("Bản vẽ thi công, cao độ cắt đầu cọc theo thiết kế", BM), ("Biện pháp thi công cắt, đập đầu cọc", BM),
        ("Dây ngăn, bảng hiệu, lối lên xuống", BM), ("Thiết bị cắt, đập và bảo hộ lao động", BM),
        ("Cọc đã ép xong và đã được nghiệm thu ép cọc", "Kiểm tra hồ sơ"), ("Hố móng đúng cao độ, khô ráo, thoát nước", BM),
        ("Cao độ đầu cọc sau khi cắt", "Bằng máy thủy bình"), ("Mặt cắt đầu cọc phẳng, vuông góc trục, không nứt vỡ", BM),
        ("Chiều dài thép chờ, thép lộ ra theo thiết kế (nếu có)", "Bằng mắt thường + thước thép"),
        ("Số lượng cọc cắt đối chiếu bảng tổng hợp ép cọc", "Kiểm tra hồ sơ"),
        ("Thu gom, bốc xúc, vận chuyển phế thải; vệ sinh hố móng", BM),
        ("Gởi trục kiểm tra", "Bằng máy toàn đạc"), ("Định vị cọc trên mặt bằng", "Bằng máy toàn đạc"),
        ("Cao độ đầu cọc sau khi cắt", "Bằng máy thủy bình")]),
}
# Bố cục "Thép cột": (hàng nhóm, tên nhóm, các hàng mục)
LAYOUT = [(22, "Công tác chuẩn bị", [23, 24]), (25, "An toàn lao động", [26, 27]), (28, "Điều kiện thi công", [29, 30]),
          (31, "Nghiệm thu", [32, 33, 34, 35, 36]), (37, "Trắc đạc", [38, 39, 40])]

wb = openpyxl.load_workbook(XLSX)
for old in [n for n in wb.sheetnames if n in NEW]:
    del wb[old]
src = wb["Thép cột"]
imgs = []
for im in src._images:  # đọc ảnh một lần rồi nạp lại để còn lưu được
    data = im._data()
    im.ref = io.BytesIO(data)
    imgs.append((data, copy.deepcopy(im.anchor), im.width, im.height))
cl = json.loads((ROOT / "checklists.json").read_text(encoding="utf-8"))
for name, (title, items) in NEW.items():
    ws = wb.copy_worksheet(src)
    ws.title = name
    for data, anchor, w, h in imgs:
        new = Image(io.BytesIO(data))
        new.anchor, new.width, new.height = copy.deepcopy(anchor), w, h
        ws.add_image(new)
    ws.print_area = src.print_area
    ws["A2"] = title
    ws["D4"] = f"Đối tượng kiểm tra: {name}"
    muc, k = [], 0
    for _, nhom, rows in LAYOUT:
        for r in rows:
            nd, pp = items[k]
            k += 1
            ws.cell(r, 2).value, ws.cell(r, 4).value = nd, pp
            if len(nd) > 42:
                ws.row_dimensions[r].height = 30
            muc.append({"nhom": nhom, "noi_dung": nd, "phuong_phap": pp})
    cl[f"XD|{name}"] = {"nhom_file": "XD", "sheet": name, "tieu_de": " ".join(title.split()), "doi_tuong": name, "tu_soan": True, "muc": muc}
wb.save(XLSX)
(ROOT / "checklists.json").write_text(json.dumps(cl, ensure_ascii=False, indent=1), encoding="utf-8")
print("ok", list(NEW))
