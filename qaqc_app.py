"""Quản lý hồ sơ chất lượng công trình xây dựng (QA/QC) - Streamlit.

Nguyên tắc: kết quả kiểm tra / kết luận do người kiểm tra thực tế nhập,
ứng dụng không tự điền "Đạt".
"""
import io
import json
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import qaqc_docx
import qaqc_xlsx

st.set_page_config(page_title="Hồ sơ chất lượng xây dựng", page_icon=":building_construction:", layout="wide")

DATA_DIR = Path(__file__).parent / "data" / "qaqc"
DATA_FILE = DATA_DIR / "records.json"

TYPES = {
    "VL": "Nghiệm thu vật liệu đầu vào",
    "CV": "Nghiệm thu công việc xây dựng",
    "GD": "Nghiệm thu giai đoạn",
    "NK": "Nhật ký thi công",
}
STATUS = ["Chưa kiểm tra", "Đạt", "Không đạt - làm lại", "Đạt có điều kiện"]
DEFAULT_CHECKLIST = {
    "VL": ["Có chứng chỉ xuất xưởng / CO-CQ", "Có kết quả thí nghiệm đầu vào", "Đúng chủng loại, quy cách theo thiết kế", "Đã lấy mẫu thí nghiệm (biên bản LM)"],
    "CV": ["Có bản vẽ thi công được duyệt", "Vật liệu đã được nghiệm thu", "Đúng vị trí, kích thước, cao độ", "Biện pháp thi công / an toàn đã duyệt", "Có kết quả đo đạc, thí nghiệm hiện trường"],
    "GD": ["Các công việc thành phần đã nghiệm thu", "Hồ sơ hoàn công giai đoạn", "Kết quả thí nghiệm đạt yêu cầu"],
    "NK": [],
}


def load():
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return {"project": {"ten_cong_trinh": "", "so_hop_dong": "", "chu_dau_tu": "", "tu_van_giam_sat": "", "nha_thau": "", "dia_diem": "", "goi_thau": ""}, "records": []}


@st.cache_data
def load_checklists():
    f = DATA_DIR / "checklists.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}


def save(db):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")


def next_number(db, t):
    n = sum(1 for r in db["records"] if r["type"] == t) + 1
    return f"{t} {n:03d}"


db = load()
st.title(":building_construction: Quản lý hồ sơ chất lượng xây dựng")
st.caption("Hồ sơ pháp lý: chỉ ghi số liệu, kết quả và chữ ký thực tế. Ứng dụng không tự điền kết luận 'Đạt'.")

with st.sidebar:
    st.header("Thông tin dự án")
    for k, label in [("ten_cong_trinh", "Tên công trình"), ("so_hop_dong", "Số hợp đồng"), ("chu_dau_tu", "Chủ đầu tư"),
                     ("tu_van_giam_sat", "Tư vấn giám sát"), ("nha_thau", "Nhà thầu"), ("dia_diem", "Địa điểm"), ("goi_thau", "Gói thầu")]:
        db["project"][k] = st.text_input(label, db["project"].get(k, ""))
    if st.button("Lưu thông tin dự án"):
        save(db)
        st.success("Đã lưu")

tab_new, tab_list, tab_log, tab_exp = st.tabs(["Lập biên bản", "Danh mục hồ sơ", "Nhật ký thi công", "Xuất hồ sơ"])

with tab_new:
    t = st.selectbox("Loại biên bản", [k for k in TYPES if k != "NK"], format_func=lambda k: f"{k} - {TYPES[k]}")
    cls = load_checklists()
    mau = ""
    if t in ("CV", "GD") and cls:
        mau = st.selectbox("Mẫu checklist theo công việc (từ file mẫu XD / MEP)", [""] + sorted(cls, key=lambda k: (cls[k]["nhom_file"], cls[k]["doi_tuong"])),
                           format_func=lambda k: "(checklist chung)" if not k else f"[{cls[k]['nhom_file']}] {cls[k]['doi_tuong'] or cls[k]['sheet']}")
    with st.form("new_record", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        so = c1.text_input("Số biên bản", next_number(db, t))
        ngay = c2.date_input("Ngày nghiệm thu (thực tế)", date.today())
        vitri = c3.text_input("Vị trí / phân đoạn")
        hangmuc = st.text_input("Hạng mục / công việc / vật liệu")
        c4, c5 = st.columns(2)
        dvt = c4.text_input("Đơn vị tính")
        kl = c5.number_input("Khối lượng", min_value=0.0, step=0.1)
        st.markdown("**Checklist kiểm tra** (chọn kết quả sau khi kiểm tra thực tế)")
        if mau:
            rows = cls[mau]["muc"]
            cl = pd.DataFrame({"Nhóm": [m["nhom"] for m in rows], "Nội dung": [m["noi_dung"] for m in rows],
                               "Phương pháp": [m["phuong_phap"] for m in rows], "Kết quả": "Chưa kiểm tra", "Ghi chú": ""})
        else:
            cl = pd.DataFrame({"Nhóm": "", "Nội dung": DEFAULT_CHECKLIST[t], "Phương pháp": "", "Kết quả": "Chưa kiểm tra", "Ghi chú": ""})
        cl = st.data_editor(cl, num_rows="dynamic", use_container_width=True,
                            column_config={"Kết quả": st.column_config.SelectboxColumn(options=STATUS)})
        lay_mau = st.checkbox("Cần lấy mẫu thí nghiệm (tạo biên bản lấy mẫu LM)") if t == "VL" else False
        kl_chung = st.selectbox("Kết luận chung", ["Chưa có kết luận"] + STATUS[1:])
        if st.form_submit_button("Tạo biên bản"):
            if any(r["so"] == so for r in db["records"]):
                st.error("Số biên bản đã tồn tại")
            else:
                db["records"].append({"type": t, "so": so, "ngay": str(ngay), "vi_tri": vitri, "hang_muc": hangmuc,
                                      "dvt": dvt, "khoi_luong": kl, "checklist": cl.fillna("").to_dict("records"),
                                      "ket_luan": kl_chung, "mau_checklist": mau, "lay_mau": lay_mau})
                save(db)
                st.success(f"Đã tạo {so}")

with tab_list:
    recs = [r for r in db["records"] if r["type"] != "NK"]
    if not recs:
        st.info("Chưa có biên bản.")
    else:
        df = pd.DataFrame(recs)
        f = st.multiselect("Lọc loại", list(TYPES), default=[k for k in TYPES if k != "NK"])
        df = df[df["type"].isin(f)].sort_values(["ngay", "so"])
        df["chua_kiem_tra"] = df["checklist"].apply(lambda c: sum(1 for i in c if i["Kết quả"] == "Chưa kiểm tra"))
        st.dataframe(df[["so", "ngay", "hang_muc", "vi_tri", "dvt", "khoi_luong", "ket_luan", "chua_kiem_tra"]], use_container_width=True)
        a, b, c = st.columns(3)
        a.metric("Số biên bản", len(df))
        b.metric("Chưa có kết luận", int((df["ket_luan"] == "Chưa có kết luận").sum()))
        c.metric("Mục checklist chưa kiểm tra", int(df["chua_kiem_tra"].sum()))
        xoa = st.selectbox("Xóa biên bản", [""] + list(df["so"]))
        if xoa and st.button("Xóa"):
            db["records"] = [r for r in db["records"] if r["so"] != xoa]
            save(db)
            st.rerun()

with tab_log:
    st.markdown("Nhật ký theo TCVN 4055. Thời tiết/nhân lực/máy phải là **ghi nhận thực tế** trong ngày.")
    with st.form("nk", clear_on_submit=True):
        d = st.date_input("Ngày", date.today())
        cviec = st.text_area("Công việc thi công")
        c1, c2, c3 = st.columns(3)
        tt = c1.text_input("Thời tiết")
        nl = c2.number_input("Nhân lực (người)", min_value=0, step=1)
        may = c3.text_input("Máy móc, thiết bị")
        gc = st.text_area("Ghi chú / sự cố / biên bản nghiệm thu trong ngày")
        if st.form_submit_button("Ghi nhật ký"):
            db["records"].append({"type": "NK", "so": f"NK {d}", "ngay": str(d), "cong_viec": cviec, "thoi_tiet": tt,
                                  "nhan_luc": nl, "may": may, "ghi_chu": gc})
            save(db)
            st.success("Đã ghi")
    nk = [r for r in db["records"] if r["type"] == "NK"]
    if nk:
        st.dataframe(pd.DataFrame(nk).sort_values("ngay")[["ngay", "cong_viec", "thoi_tiet", "nhan_luc", "may", "ghi_chu"]], use_container_width=True)

with tab_exp:
    st.markdown("Xuất danh mục hồ sơ, checklist và nhật ký ra Excel.")
    if db["records"]:
        buf = io.BytesIO()
        recs = [r for r in db["records"] if r["type"] != "NK"]
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            pd.DataFrame([{"Số": r["so"], "Loại": TYPES[r["type"]], "Ngày": r["ngay"], "Hạng mục": r["hang_muc"], "Vị trí": r["vi_tri"],
                           "ĐVT": r["dvt"], "Khối lượng": r["khoi_luong"], "Kết luận": r["ket_luan"]} for r in recs]).to_excel(w, sheet_name="Danh muc", index=False)
            pd.DataFrame([{"Số": r["so"], **i} for r in recs for i in r["checklist"]]).to_excel(w, sheet_name="Checklist", index=False)
            pd.DataFrame([r for r in db["records"] if r["type"] == "NK"]).to_excel(w, sheet_name="Nhat ky", index=False)
        st.download_button("Tải Danh_muc_ho_so.xlsx", buf.getvalue(), "Danh_muc_ho_so.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("---")
        st.markdown("**Xuất biên bản Word** theo form mẫu của TVGS (chỉ điền dữ liệu đã nhập; chỗ chưa có giữ dấu chấm để điền tay).")
        sel = st.multiselect("Chọn biên bản", [r["so"] for r in recs if r["type"] in qaqc_docx.FORMS])
        if sel:
            kinds = st.multiselect("Loại biểu mẫu", ["bb", "yc", "lm"], default=["bb"],
                                   format_func=lambda k: {"bb": "Biên bản nghiệm thu", "yc": "Phiếu yêu cầu nghiệm thu", "lm": "Biên bản lấy mẫu (VL có đánh dấu lấy mẫu)"}[k])
            chon = qaqc_docx.expand([r for r in recs if r["so"] in sel], kinds)
            st.download_button("Tải Bien_ban_nghiem_thu.docx", qaqc_docx.build_docx(chon, db["project"]), "Bien_ban_nghiem_thu.docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.caption("Hiện hỗ trợ form VL (/CV/VL-TB) và CV (/CV).")
        st.markdown("**Xuất checklist Excel** đúng form gốc (chỉ cho biên bản đã chọn mẫu checklist). Mục 'Chưa kiểm tra' để ô trống; cột L2 của TVGS luôn để trống.")
        co_mau = [r for r in recs if r.get("mau_checklist")]
        sel2 = st.multiselect("Chọn biên bản xuất checklist", [r["so"] for r in co_mau])
        if sel2:
            chon2 = sorted([r for r in co_mau if r["so"] in sel2], key=lambda r: (r["ngay"], r["so"]))
            if len(chon2) == 1:
                fn, data = qaqc_xlsx.build_checklist_xlsx(chon2[0], db["project"])
                st.download_button(f"Tải {fn}", data, fn, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.download_button("Tải Checklist.zip", qaqc_xlsx.build_checklists_zip(chon2, db["project"]), "Checklist.zip", "application/zip")
    else:
        st.info("Chưa có dữ liệu.")
