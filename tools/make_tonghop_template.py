"""Trích sheet "tổng hợp công tác ép cọc" từ Checklist_XD.xlsx thành file mẫu nhỏ."""
import sys
from pathlib import Path

import openpyxl

wb = openpyxl.load_workbook(sys.argv[1])
for n in list(wb.sheetnames):
    if n != "tổng hợp công tác ép cọc":
        del wb[n]
wb._external_links = []
wb.save(Path(__file__).parent.parent / "data/qaqc/templates/tonghop_ep_coc.xlsx")
print("ok")
