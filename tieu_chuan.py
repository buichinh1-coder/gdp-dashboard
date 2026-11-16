"""Tiêu chuẩn áp dụng theo từng loại công việc / vật liệu (đã tra cứu, xem sheet Tieu_chuan trong file 07).

Chỉ đưa vào biểu mẫu các tiêu chuẩn đã xác nhận được tên và số hiệu. Tiêu chuẩn chưa chắc hiệu lực hoặc chưa
xác minh được nằm ở danh sách TC_CAN_XAC_NHAN, không tự đưa vào biên bản.
"""
DAT = ["TCVN 4447:2012"]
BT = ["TCVN 4453:1995"]
TC_CONG_VIEC = {  # theo Mục HĐ của lịch nghiệm thu
    "15": DAT, "16": DAT, "05": DAT, "06": DAT, "07": DAT,
    "04": ["TCVN 4447:2012", "TCVN 9361:2012"],
    "02": ["TCVN 9394:2012"],
    "03": ["TCVN 9394:2012", "TCVN 9361:2012"],
    "08": BT, "10": BT, "11": BT, "12": BT, "14": BT,
    "13": ["TCVN 4453:1995", "TCVN 1651-2:2018", "QCVN 07:2019/BKHCN"],
    "09": ["TCVN 4453:1995", "TCVN 3105:2022", "TCVN 3106:2022", "TCVN 3118:2022", "TCVN 8828:2011"],
    "17": [],
}
TC_VAT_LIEU = {
    "Cọc bê tông": ["TCVN 9394:2012", "TCVN 3118:2022", "TCVN 7570:2006", "TCVN 4506:2012", "TCVN 1651-2:2018"],
    "Cốt thép": ["TCVN 1651-2:2018", "QCVN 07:2019/BKHCN"],
    "M300": ["TCVN 4453:1995", "TCVN 3105:2022", "TCVN 3106:2022", "TCVN 3118:2022", "TCVN 3116:2022", "TCVN 8828:2011"],
}
CHUNG = "Hồ sơ thiết kế, chỉ dẫn kỹ thuật của hợp đồng"

# (số hiệu, tên, tình trạng tra cứu, nguồn)
TC_BANG = [
    ("TCVN 9394:2012", "Đóng và ép cọc - Thi công và nghiệm thu", "Đã xác nhận số hiệu và tên", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-9394-2012-dong-va-ep-coc-Thi-cong-va-nghiem-thu-906973.aspx"),
    ("TCVN 9393:2012", "Cọc - Phương pháp thí nghiệm bằng tải trọng tĩnh ép dọc trục", "Đã xác nhận; chỉ dùng nếu thiết kế yêu cầu thử tải", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-9394-2012-dong-va-ep-coc-Thi-cong-va-nghiem-thu-906973.aspx"),
    ("TCVN 4447:2012", "Công tác đất - Thi công và nghiệm thu", "Đã xác nhận, còn dẫn chiếu trong tiêu chuẩn 2026", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-4447-2012-Cong-tac-dat-Thi-cong-va-nghiem-thu-906964.aspx"),
    ("TCVN 9361:2012", "Công tác nền móng - Thi công và nghiệm thu", "Đã xác nhận số hiệu và tên", "https://luatvietnam.vn/xay-dung/tieu-chuan-tcvn-9361-2012-thi-cong-nghiem-thu-nen-mong-160861-d3.html"),
    ("TCVN 4453:1995", "Kết cấu bê tông và bê tông cốt thép toàn khối - Quy phạm thi công và nghiệm thu", "Còn được dẫn chiếu; đang có dự thảo TCVN 4453:202x thay thế. CẦN xác nhận hiệu lực với CĐT/TVGS", "https://dauthau.gxd.vn/van-ban/chat-luong/tcvn-4453-1995.html"),
    ("TCVN 3105:2022", "Hỗn hợp bê tông và bê tông - Lấy mẫu, chế tạo và bảo dưỡng mẫu thử", "Đã xác nhận (QĐ 2142/QĐ-BKHCN, thay TCVN 3105:1993)", "https://luatvietnam.vn/khoa-hoc/quyet-dinh-2142-qd-bkhcn-238456-d1.html"),
    ("TCVN 3106:2022", "Hỗn hợp bê tông - Phương pháp xác định độ sụt", "Đã xác nhận (thay TCVN 3106:1993)", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-3106-2022-Hon-hop-be-tong-phuong-phap-xac-dinh-do-sut-919666.aspx"),
    ("TCVN 3118:2022", "Bê tông - Phương pháp xác định cường độ chịu nén", "Số hiệu và tên có trong QĐ 2142/QĐ-BKHCN; chưa đọc nội dung chi tiết", "https://luatvietnam.vn/khoa-hoc/quyet-dinh-2142-qd-bkhcn-238456-d1.html"),
    ("TCVN 3116:2022", "Bê tông - Phương pháp xác định độ chống thấm nước (phương pháp vết thấm)", "Đã xác nhận (thay TCVN 3116:1993)", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-3116-2022-Be-tong-Phuong-phap-xac-dinh-do-chong-tham-nuoc-vet-tham-919665.aspx"),
    ("TCVN 8828:2011", "Bê tông - Yêu cầu bảo dưỡng ẩm tự nhiên", "Đã xác nhận", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-8828-2011-Be-tong-Yeu-cau-bao-duong-am-tu-nhien-907381.aspx"),
    ("TCVN 1651-2:2018", "Thép cốt bê tông - Phần 2: Thép thanh vằn", "Đã xác nhận (thay TCVN 1651-2:2008)", "https://thuvienphapluat.vn/TCVN/xay-dung/tcvn-1651-2-2018-thep-cot-be-tong-phan-2-thep-thanh-van-917774.aspx"),
    ("QCVN 07:2019/BKHCN", "Quy chuẩn kỹ thuật quốc gia về thép làm cốt bê tông", "Đã xác nhận, hiệu lực từ 01/01/2020", "https://thuvienphapluat.vn/TCVN/Xay-dung/QCVN-7-2019-BKHCN-Thep-lam-cot-be-tong-918504.aspx"),
    ("TCVN 7570:2006", "Cốt liệu cho bê tông và vữa - Yêu cầu kỹ thuật", "Đã xác nhận", "https://luatminhkhue.vn/van-ban/tieu-chuan-viet-nam-tcvn-7570-2006-ve-cot-lieu-cho-be-tong-va-vua-yeu-cau-ky-thuat-do-bo-khoa-hoc-va-cong-nghe-ban-hanh.aspx"),
    ("TCVN 4506:2012", "Nước cho bê tông và vữa - Yêu cầu kỹ thuật", "Đã xác nhận số hiệu; nguồn ghi tên hơi khác nhau", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-10796-2015-Cat-min-cho-be-tong-va-vua-917728.aspx"),
]
TC_CAN_XAC_NHAN = [
    ("TCVN 9340:2012", "Hỗn hợp bê tông trộn sẵn", "Tra cứu có thông tin TCVN 14586:2025 (Hỗn hợp bê tông - Yêu cầu kỹ thuật) thay thế. Chưa đưa vào biên bản; xác nhận tiêu chuẩn bê tông thương phẩm trong hợp đồng.", "https://thuvienphapluat.vn/TCVN/Xay-dung/TCVN-14586-2025-Hon-Hop-Be-Tong-Yeu-Cau-Ky-Thuat-922245.aspx"),
    ("TCVN 6067:2018", "Xi măng poóc lăng bền sunfat", "Nguồn ghi năm khác nhau (2018/2004). Chưa đưa vào biên bản; đối chiếu chỉ dẫn kỹ thuật.", ""),
    ("Tiêu chuẩn sản xuất cọc BTCT 300x300 tự sản xuất", "", "Tra cứu không xác nhận được số hiệu áp dụng cho cọc BTCT vuông thường. Theo chỉ dẫn kỹ thuật/thiết kế của hợp đồng.", ""),
    ("Nghị định về quản lý chất lượng thi công xây dựng", "", "Nghị định 06/2021/NĐ-CP đã được sửa đổi nhiều lần; có thông tin về Nghị định 207/2026/NĐ-CP (15/06/2026) và văn bản hợp nhất 63/2026/VBHN-BXD. Chưa viện dẫn trong biên bản; xác nhận văn bản hiện hành tại thời điểm ký hợp đồng 12/08/2026.", "https://luatvietnam.vn/linh-vuc-khac/bang-so-sanh-nghi-dinh-207-2026-nd-cp-va-nghi-dinh-06-2021-nd-cp-ve-quan-ly-chat-luong-thi-cong-xay-dung-883-109774-article.html"),
]


def tieu_chuan_cv(muc):
    return TC_CONG_VIEC.get(muc, [])


def tieu_chuan_vl(hang_muc):
    for k, v in TC_VAT_LIEU.items():
        if k in hang_muc:
            return v
    return []
