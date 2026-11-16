"""Rút gọn file mẫu checklist: chỉ giữ sheet checklist, bỏ liên kết ngoài.
Dùng: python tools/strip_checklist_templates.py <XD.xlsx> <MEP.xlsx>"""
import io
import json
import sys
from pathlib import Path

import openpyxl

def _nho(img):
    """Đọc ảnh một lần, nạp lại vào BytesIO để còn lưu được; True nếu là ảnh nhỏ (logo)."""
    data = img._data()
    img.ref = io.BytesIO(data)
    return len(data) < 50000


root = Path(__file__).parent.parent / "data" / "qaqc"
cl = json.loads((root / "checklists.json").read_text(encoding="utf-8"))
for group, src in zip(("XD", "MEP"), sys.argv[1:3]):
    keep = {v["sheet"] for v in cl.values() if v["nhom_file"] == group}
    wb = openpyxl.load_workbook(src)
    for n in list(wb.sheetnames):
        if n not in keep:
            del wb[n]
    for ws in wb.worksheets:  # bỏ sơ đồ/ảnh lớn của dự án khác, giữ logo
        ws._images = [i for i in ws._images if _nho(i)]
    wb._external_links = []
    wb.save(root / "templates" / f"checklist_{group}.xlsx")
    print(group, len(wb.sheetnames))
