# Ghi chú bộ hồ sơ nghiệm thu – Thi công kè chắn đất, cọc 88–126

Bộ hồ sơ dựng từ `Lich_nghiem_thu_ke_BT_coc_88_126.xlsx`. **Đây là bản khung cần hoàn thiện.**

## Đã điền / để trống
- Đã điền: số biên bản, hạng mục, vị trí, khối lượng, ngày nghiệm thu.
- **Ngày nghiệm thu** lấy theo ngày kế hoạch, theo xác nhận của người lập hồ sơ rằng đó là ngày thực tế.
- Thông tin dự án (theo hợp đồng 12.08/2026/HĐKT/ĐGC-GL): CĐT Công ty TNHH Đỗ Gia Capital; nhà thầu Công ty TNHH Giang Long;
  gói thầu "Thi công kè chắn đất bê tông đổ tại chỗ tuyến 3 đoạn từ cọc 88 đến cọc 126"; địa điểm Khu đô thị tại khu 10B,
  phường Quang Hanh, TP Cẩm Phả, Quảng Ninh.
- Để trống: kết quả kiểm tra (mọi mục "Chưa kiểm tra"), kết luận, tên và chữ ký, số bản vẽ, số phiếu yêu cầu.
  Cột L2 (TVGS) trong checklist để trống.
- Giữ từ form mẫu: tên TVGS (Apave) và logo CĐT/TVGS trong checklist.
- Logo nhà thầu trong checklist là logo tự tạo (`data/qaqc/templates/logo_giang_long.png`); thay bằng logo chính thức nếu có.

## Giả định cần xác nhận
1. Nghiệm thu vật liệu đầu vào (VL):
   - Cọc BT300x300: 1 lần cho cả 3.372 m (VL 001), ngày 24/08/2026, ngày trước khi ép cọc đầu tiên.
     Phải đủ toàn bộ cọc và đạt R7 trước ngày này. Nếu cọc phân đoạn 2–4 sản xuất sau thì cần tách thêm biên bản.
   - Thép: 1 lần cho 23,02 tấn, ngày 04/09/2026 (1 ngày trước ngày bắt đầu thép phân đoạn 1). Ngày này do người lập chọn, lịch không có.
   - Bê tông thương phẩm M300: 1 biên bản cho mỗi đợt đổ (4 đợt), ngày = ngày bắt đầu đổ. Chưa có VL cho bê tông lót M100,
     đá xô, vật liệu đắp, vật tư khe lún.
2. Các dòng "cung cấp cọc" theo phân đoạn trong lịch được gộp vào VL 001, nên không còn là biên bản công việc.
3. Mẫu checklist là mẫu gần nhất trong file XD:
   - đắp đê quai, đắp hoàn trả, đá xô -> "San lấp";
   - nạo vét, đào móng -> "Đào đất";
   - ván khuôn -> "Ván khuôn dầm sàn";
   - thép -> "Thép dầm sàn";
   - đổ bê tông -> "Đổ bê tông";
   - BT lót -> "Bê tông lót";
   - khe lún -> "Khe lún".
4. Bê tông phụ gia R7 đạt 100% cường độ: cần kết quả nén mẫu R7 làm căn cứ khi đắp đất lên tường chắn.

## Bổ sung lần này
- Checklist **ép cọc** và **cắt đầu cọc** (8 file): file mẫu gốc không có hai checklist này, nên tôi soạn theo bố cục
  form "Thép cột" (cùng khung, logo, cột L1/L2). Nội dung chỉ liệt kê hạng mục cần kiểm tra, không ghi dung sai bằng số.
  Cần TVGS và kỹ sư rà lại với chỉ dẫn kỹ thuật và TCVN 9394:2012.
- `05_Phieu_yeu_cau_nghiem_thu/`: phiếu yêu cầu nghiệm thu (YC/VL-TB cho vật liệu, YC/CV cho công việc), cùng số với biên bản.
  **Ngày lập phiếu để trống**; ô "Thời gian nghiệm thu" đã điền ngày nghiệm thu.
- `06_Bien_ban_lay_mau.docx`: 6 biên bản lấy mẫu thí nghiệm (cọc, thép, 4 đợt bê tông M300). Tiêu chuẩn áp dụng, khối lượng và
  quy cách lấy mẫu, thành phần tham gia để trống cho người lấy mẫu điền.
  Ngày lấy mẫu tạm lấy bằng ngày nghiệm thu vật liệu (cọc 24/08, thép 04/09, bê tông = ngày bắt đầu đổ); cần sửa theo ngày lấy mẫu thực tế.
- `07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx`: danh sách kết quả thí nghiệm và hồ sơ cần có (cột trạng thái "Cần bổ sung"),
  kèm sheet mốc nén R7 từng đợt đổ so với ngày bắt đầu đắp. Tiêu chuẩn ghi để tham khảo, cần đối chiếu chỉ dẫn kỹ thuật.
  Nhóm D là đề xuất, chưa được xác nhận.

## Chưa có checklist (5 biên bản)
Vận chuyển đất (4 biên bản) và biện pháp thi công (1 biên bản): file mẫu chưa có checklist cho các công việc này.

## Chưa lập (app chưa hỗ trợ form)
Nghiệm thu hoàn thành toàn gói (CV/CT), bàn giao, nhật ký thi công, biên bản kiểm tra chế tạo cọc, trắc đạc cọc.

## Cần rà lại trong lịch
- Cắt đầu cọc 123–124 cọc/ngày.
- Đào móng xong trước khi ép cọc bắt đầu.
- 7 mốc nghiệm thu rơi vào Chủ nhật.
- Hàm lượng thép khoảng 42 kg/m³ bê tông.
- Cuối gói: nghiệm thu hoàn thành, bàn giao và hoàn công cùng ngày 27/10.
