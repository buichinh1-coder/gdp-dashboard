"""Trích cấu trúc checklist từ file mẫu Checklist_XD.xlsx / Checklist_MEP.xlsx -> data/qaqc/checklists.json.

Chỉ lấy nội dung công việc kiểm tra + phương pháp. KHÔNG lấy ô đã tick sẵn trong mẫu.
Dùng: python tools/extract_checklists.py <XD.xlsx> <MEP.xlsx>
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

SRC = {"XD": sys.argv[1], "MEP": sys.argv[2]}
out = {}
for group, path in SRC.items():
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        head = next((r for r in range(1, 40) if str(ws.cell(r, 2).value or "").strip().startswith("Công việc kiểm tra")), None)
        if not head:
            continue
        title = " ".join(str(ws["A2"].value or ws.title).split())
        doi_tuong = re.sub(r"[….\s]{3,}$", "", re.sub(r"^.*?:\s*[-–]?\s*", "", str(ws["D4"].value or ""))).strip()
        items, section = [], ""
        for r in range(head + 2, ws.max_row + 1):
            a = str(ws.cell(r, 1).value or "").strip()
            b = " ".join(str(ws.cell(r, 2).value or "").split())
            d = " ".join(str(ws.cell(r, 4).value or "").split())
            if a.startswith(("Ý kiến", "KẾT LUẬN")):
                break
            if not b:
                continue
            if a.isdigit():
                items.append({"nhom": section, "noi_dung": b, "phuong_phap": d})
            elif a:
                section = b
        if items:
            out[f"{group}|{ws.title}"] = {"nhom_file": group, "sheet": ws.title, "tieu_de": title,
                                                  "doi_tuong": doi_tuong, "muc": items}
Path(__file__).parent.parent.joinpath("data/qaqc/checklists.json").write_text(
    json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
print(len(out), "mẫu checklist")
