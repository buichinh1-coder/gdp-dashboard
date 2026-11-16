# Ghi chú bộ hồ sơ nghiệm thu – Thi công kè chắn đất, cọc 88–126

Bộ hồ sơ dựng từ `Lich_nghiem_thu_ke_BT_coc_88_126.xlsx`. **Đây là bản khung cần hoàn thiện.**

## Quyển hồ sơ hoàn chỉnh
- `00_HO_SO_HOAN_CHINH.pdf` (288 trang, có bookmark): bìa, trang xác nhận, danh mục hồ sơ có số trang, tờ phân cách các phần A-D,
  phiếu yêu cầu, biên bản, checklist và bảng tổng hợp thi công cọc, đánh số trang "Trang x/288" ở chân trang.
  Phần D là tờ phân cách để đặt chứng chỉ, kết quả thí nghiệm, bản vẽ hoàn công, nhật ký (chưa có nên chưa kèm).
- `00_Bia_ho_so.docx`, `00_Muc_luc_ho_so.docx`, `00_Danh_muc_ho_so.xlsx`: bản sửa được của bìa và danh mục.
- Dựng lại sau khi chỉnh các file: `python tools/assemble_ho_so.py ho_so` (số trang mục lục tự cập nhật).

## Đã điền / để trống
- Đã điền: số biên bản, hạng mục, vị trí, khối lượng, ngày nghiệm thu.
- **Ngày nghiệm thu** lấy theo ngày kế hoạch, theo xác nhận của người lập hồ sơ rằng đó là ngày thực tế.
- Thông tin dự án (theo hợp đồng 12.08/2026/HĐKT/ĐGC-GL): CĐT Công ty TNHH Đỗ Gia Capital; nhà thầu Công ty TNHH Giang Long;
  gói thầu "Thi công kè chắn đất bê tông đổ tại chỗ tuyến 3 đoạn từ cọc 88 đến cọc 126"; địa điểm Khu đô thị tại khu 10B,
  phường Quang Hanh, TP Cẩm Phả, Quảng Ninh.
- Người ký hồ sơ: Chỉ huy trưởng Giang Long Nguyễn Mạnh Tùng và cán bộ kỹ thuật Apave Nguyễn Thế Anh. Tên in sẵn ở mục
  thành phần tham gia, khối chữ ký của biên bản, phiếu yêu cầu, biên bản lấy mẫu và cột "Họ tên" của checklist.
  Giám sát trưởng Nguyễn Huy Hà chỉ có tên ở mục thành phần tham gia (không in ở ô ký). Chữ ký để trống.
- Để trống: kết quả kiểm tra (mọi mục "Chưa kiểm tra"), kết luận, tên và chữ ký, số bản vẽ, số phiếu yêu cầu.
  Cột L2 (TVGS) trong checklist để trống.
- Giữ từ form mẫu: tên TVGS (Apave) và logo CĐT/TVGS trong checklist.
- Ô logo nhà thầu trong checklist và bìa chỉ ghi tên "CÔNG TY TNHH GIANG LONG" (chữ, không dùng hình logo), theo yêu cầu.

## Tiêu chuẩn áp dụng (đã tra cứu trực tuyến)
- Dòng "Tiêu chuẩn, quy phạm xây dựng được áp dụng" của biên bản nghiệm thu và mục 5 của biên bản lấy mẫu đã điền theo từng
  loại công việc, chỉ gồm tiêu chuẩn xác nhận được số hiệu và tên (sheet `Tieu_chuan_da_dung` trong file 07).
- Chưa đưa vào biên bản, cần bạn xác nhận (sheet `Tieu_chuan_can_xac_nhan`): TCVN 4453:1995 còn hiệu lực hay đã có bản thay thế,
  tiêu chuẩn bê tông thương phẩm (TCVN 9340:2012 hay TCVN 14586:2025), xi măng bền sunfat, tiêu chuẩn sản xuất cọc BTCT 300x300,
  và nghị định quản lý chất lượng hiện hành (có thông tin về Nghị định 207/2026/NĐ-CP).

## Quy ước trình bày (theo yêu cầu)
- Khối lượng chỉ ghi ở biên bản nghiệm thu vật liệu đầu vào, biên bản lấy mẫu và bảng tổng hợp cọc; biên bản/phiếu yêu cầu nghiệm thu công việc không ghi khối lượng.
- Độ chặt **K95** chỉ áp dụng công tác đắp đất: đắp đê quai và đắp hoàn trả hố móng (ghi trong tên công việc và danh sách thí nghiệm).
  Không áp dụng cho đá xô, đào, nạo vét. Giả định đê quai là đất đắp; cần xác nhận.
- Biên bản lấy mẫu bê tông M300 lấy ngày đổ bê tông (bắt đầu - kết thúc đổ của từng đợt) làm ngày lấy mẫu.
- Khối chữ ký: tên in sát dưới ô ký (2-4 dòng), hai ô căn giữa đối xứng.

## Thí nghiệm đã chốt (theo yêu cầu)
- Phòng thí nghiệm LAS-XD 1191 (điền vào biên bản lấy mẫu và danh sách thí nghiệm). Ngày phiếu kết quả chốt bằng ngày thực hiện
  thí nghiệm; số phiếu và kết quả để trống.
- Thép chỉ thí nghiệm kéo (không uốn). Bê tông M300 chỉ nén R7 và R28 (không độ sụt, không thấm nước vì chỉ dẫn kỹ thuật không yêu cầu).
- Bỏ thí nghiệm đá (đá xô chỉ nghiệm thu vật liệu đầu vào) và nén tĩnh cọc. **Giữ** thí nghiệm đầm nén tiêu chuẩn của đất đắp (đê quai 17/08, hoàn trả 20/09)
  làm cơ sở cho độ chặt K95, làm tại hiện trường theo từng ngày đắp.

## Khối lượng và quy cách lấy mẫu bê tông (biên bản lấy mẫu)
- Bê tông M300 mỗi đợt đổ: tần suất **1 tổ mẫu cho mỗi 20 m³** (TCVN 4453:1995, nhóm kết cấu khung, cột, dầm, sàn; làm tròn lên) = **7 lần lấy mẫu**
  cho mỗi đợt 136,94 m³; mỗi lần đúc 2 tổ (1 tổ nén R7, 1 tổ nén R28), mỗi tổ 3 viên lập phương 150x150x150 mm (TCVN 3105:2022): 14 tổ, 42 viên mỗi đợt.
- Bê tông cọc (VL 001): khối lượng 3.372 m cọc 300x300 = 303,48 m³; tần suất 1 tổ/20 m³ → **16 tổ nén R7 (48 viên)**, lấy trong thời gian đúc cọc (14/08 - 24/08/2026).
  Tần suất cho cọc đúc sẵn lấy theo cách tính bê tông thông thường vì không tìm được quy định riêng cho cọc đúc sẵn.
- Cốt thép (VL 002): 23,02 tấn, dưới 50 tấn mỗi lô; mỗi đường kính, mỗi mác 1 tổ gồm 3 mẫu thử kéo (TCVN 1651-2:2018). Số tổ bằng số đường kính thép dùng,
  lịch không có nên ghi theo quy tắc, chưa ghi con số.
- **Cần xác nhận:** không mở được toàn văn TCVN 4453:1995 và TCVN 3105:2022 (mạng của phiên làm việc không truy cập được các trang tiêu chuẩn), nên tần suất và kích thước viên mẫu
  lấy qua tra cứu nguồn thứ cấp. Một số nguồn ghi móng lớn 100 m³/tổ; viên mẫu theo TCVN 3105:2022 phụ thuộc cỡ hạt cốt liệu. Đối chiếu chỉ dẫn kỹ thuật và Apave.

## Giả định cần xác nhận
1. Nghiệm thu vật liệu đầu vào (VL):
   - Cọc BT300x300: 1 lần cho cả 3.372 m (VL 001), ngày 24/08/2026, ngày trước khi ép cọc đầu tiên.
     Phải đủ toàn bộ cọc và đạt R7 trước ngày này. Nếu cọc phân đoạn 2–4 sản xuất sau thì cần tách thêm biên bản.
   - Thép: 1 lần cho 23,02 tấn, ngày 04/09/2026 (1 ngày trước ngày bắt đầu thép phân đoạn 1). Ngày này do người lập chọn, lịch không có.
   - Bê tông thương phẩm M300: **không nghiệm thu vật liệu riêng**. Lấy mẫu và thí nghiệm R7, R28 nằm trong nghiệm thu công việc đổ bê tông M300
     (4 đợt đổ, mỗi đợt 1 biên bản lấy mẫu LM đi kèm biên bản đổ bê tông).
   - Bê tông lót M100: theo xác nhận, **không nghiệm thu vật liệu đầu vào, không lấy mẫu, không thí nghiệm**.
   - Đá xô: nghiệm thu vật liệu đầu vào 1 lần (VL 003, ngày 25/09/2026, 1 ngày trước ngày đắp đá xô đầu tiên), không lấy mẫu, không thí nghiệm.
   - Chưa có VL cho vật liệu đắp đất, vật tư khe lún và ống thoát nước.
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

## Gộp biên bản cho gọn
- Đào móng và vận chuyển đất đào: 1 biên bản cho mỗi phân đoạn (4 biên bản thay cho 8), 1 checklist đào đất.
- Ván khuôn bê tông lót và đổ bê tông lót: 1 biên bản cho mỗi phân đoạn (4 thay cho 8), ngày nghiệm thu lấy theo ngày đổ bê tông lót,
  kèm 2 checklist `/KT-1` (ván khuôn) và `/KT-2` (đổ bê tông lót).
- Cốt thép, ván khuôn cấu kiện, ván khuôn khóa đỉnh và khe lún: 1 biên bản "nghiệm thu trước khi đổ bê tông" cho mỗi phân đoạn (4 biên bản thay cho 16),
  vì cả 4 việc có cùng ngày nghiệm thu trong lịch (10/09, 15/09, 20/09, 25/09). Kèm 3 checklist `/KT-1` (cốt thép), `/KT-2` (ván khuôn, gồm cấu kiện và khóa đỉnh),
  `/KT-3` (khe lún).
- Nghiệm thu cọc gộp ép cọc và cắt đầu cọc như đã nêu bên dưới. Tổng số biên bản nghiệm thu công việc giảm từ 61 xuống 40.
- **Không lập nghiệm thu công việc "Biện pháp thi công trọn gói"** (không có loại nghiệm thu này), nên không còn file toàn tuyến.
- Số hiệu đánh lại theo **phân đoạn rồi theo trình tự thi công**: PĐ1 = CV 001-010, PĐ2 = 011-020, PĐ3 = 021-030, PĐ4 = 031-040. Trong mỗi phân đoạn:
  đê quai, nạo vét (70%), nghiệm thu cọc, đào móng + vận chuyển, ván khuôn + bê tông lót, thép + ván khuôn + khe lún, đổ bê tông M300, đắp hoàn trả,
  đắp đá xô, nạo vét hoàn thiện (30%). Vật liệu VL 001-003, lấy mẫu LM 001-006, giai đoạn GĐ 001. Phiếu yêu cầu cùng số với biên bản,
  checklist theo số biên bản (`/KT`, `/KT-1`...). Lưu ý: ngày nghiệm thu đổ bê tông (theo mốc R28) muộn hơn ngày đắp hoàn trả nên số hiệu và ngày không cùng thứ tự.

## Nghiệm thu cọc (gộp theo phân đoạn + giai đoạn toàn tuyến)
- Mỗi phân đoạn có **1 biên bản nghiệm thu cọc** gộp ép cọc và cắt đầu cọc (4 biên bản thay cho 8), kèm 2 checklist `/KT-1` (ép cọc)
  và `/KT-2` (cắt đầu cọc).
- Thêm **1 biên bản nghiệm thu giai đoạn cọc toàn tuyến** `GD 001` (ngày 17/09/2026, ngày nghiệm thu cọc phân đoạn cuối) kèm phiếu yêu cầu
  `001/YC/GĐ` và `03_Bang_tong_hop_thi_cong_coc.xlsx`. Bảng tổng hợp mới có số cọc, khối lượng và ngày ép theo từng phân đoạn;
  chiều dài, cao độ, tọa độ, lực ép từng cọc điền từ nhật ký ép cọc.
- Nghiệm thu cọc theo phân đoạn vẫn giữ vì các phân đoạn đầu đã làm bê tông lót, thép, ván khuôn trước khi phân đoạn cuối ép xong.
  Cần Apave chấp thuận cách gộp này.

## Rà soát trình tự công việc (lần này)
Đã áp dụng:
1. Bê tông lót M100: bỏ qua nghiệm thu vật liệu, lấy mẫu và thí nghiệm (theo xác nhận). Bộ hồ sơ chỉ có VL bê tông M300.
2. Thí nghiệm bê tông M300 làm **theo từng đợt đổ** (4 đợt), nén cả R7 và R28. Ngày nghiệm thu kết cấu bê tông giữ theo mốc R28 của lịch
   (11/10, 16/10, 21/10, 26/10). Việc đắp hoàn trả hố móng và đá xô bắt đầu trước mốc này (lúc bê tông 9 ngày tuổi), nên chỉ được
   đắp khi kết quả nén R7 đạt (theo xác nhận phụ gia R7 đạt 100%) và có chấp thuận của TVGS. Xem sheet `Moc_R7_R28` trong file 07.
3. Các biên bản cùng ngày được đánh số theo trình tự công nghệ (ví dụ ép cọc trước đào móng, cắt đầu cọc sau).
Còn để ngỏ: đào móng hoàn thành trước ngày bắt đầu ép cọc ở cả 4 phân đoạn, cần xác nhận trình tự thi công thực tế.

## Bổ sung lần này
- Checklist **ép cọc** và **cắt đầu cọc** (mỗi phân đoạn 2 file `_ep_coc`, `_cat_dau_coc`): file mẫu gốc không có hai checklist này, nên tôi soạn theo bố cục
  form "Thép cột" (cùng khung, logo, cột L1/L2). Nội dung chỉ liệt kê hạng mục cần kiểm tra, không ghi dung sai bằng số.
  Cần TVGS và kỹ sư rà lại với chỉ dẫn kỹ thuật và TCVN 9394:2012.
- `05_Phieu_yeu_cau_nghiem_thu/`: phiếu yêu cầu nghiệm thu (YC/VL-TB cho vật liệu, YC/CV cho công việc), cùng số với biên bản.
  **Ngày lập phiếu để trống**; ô "Thời gian nghiệm thu" đã điền ngày nghiệm thu.
- `06_Bien_ban_lay_mau.docx`: 6 biên bản lấy mẫu thí nghiệm (cọc, thép, 4 đợt bê tông M300). Tiêu chuẩn áp dụng, khối lượng và
  quy cách lấy mẫu, thành phần tham gia để trống cho người lấy mẫu điền.
  Ngày lấy mẫu tạm lấy bằng ngày nghiệm thu vật liệu (cọc 24/08, thép 04/09, bê tông = ngày bắt đầu đổ); cần sửa theo ngày lấy mẫu thực tế.
- `07_Danh_sach_ket_qua_thi_nghiem_can_co.xlsx`: sheet `Danh_sach` liệt kê 70 thí nghiệm cần có, mỗi dòng một thí nghiệm với ngày lấy mẫu/hiện trường
  và **ngày thực hiện thí nghiệm** cụ thể (bê tông: độ sụt cùng ngày đổ, nén R7 = ngày đổ + 7, nén R28 = ngày đổ + 28, thấm nước = ngày đổ + 28,
  tính theo từng ngày đổ; độ chặt K95 theo từng ngày đắp; cọc, thép, đầm nén, đá xô ghi hạn "chậm nhất"). Sheet `Ho_so_kem_theo` là các hồ sơ
  không phải thí nghiệm (cấp phối, phiếu giao hàng, kiểm định, nhật ký ép, trắc đạc, hoàn công). Các cột Phòng thí nghiệm đến Đạt/Không đạt để trống.
  Ngày nén tính theo ngày đổ trong lịch; nếu đổ lệch ngày thì tính lại.

## Checklist
- **Tích sẵn "Đạt" (cột L1) cho mọi mục kiểm tra theo yêu cầu của người lập hồ sơ.** Đây là bản khung: người kiểm tra thực tế phải rà từng mục,
  bỏ tick nếu mục nào không đạt hoặc chưa kiểm tra. Dòng kết luận "Đồng ý chuyển công tác tiếp theo" và cột L2 (TVGS) vẫn để trống.
- Mọi biên bản nghiệm thu công việc hiện đều có checklist (đào móng và vận chuyển đất đào dùng chung mẫu "Đào đất"; một số mẫu là mẫu gần nhất, xem phần giả định).

## Chưa lập (app chưa hỗ trợ form)
Nghiệm thu hoàn thành toàn gói (CV/CT), bàn giao, nhật ký thi công, biên bản kiểm tra chế tạo cọc, trắc đạc cọc.

## Cần rà lại trong lịch
- Cắt đầu cọc 123–124 cọc/ngày.
- Đào móng xong trước khi ép cọc bắt đầu.
- 7 mốc nghiệm thu rơi vào Chủ nhật.
- Hàm lượng thép khoảng 42 kg/m³ bê tông.
- Cuối gói: nghiệm thu hoàn thành, bàn giao và hoàn công cùng ngày 27/10.
