# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7.8 (21/10/2026 - Bản Fix Kaizen Matrix & Panel Lock)

* **Tệp Excel chính thức**: [`261008_Stock_FG_MA_BODY.xlsm`](./261008_Stock_FG_MA_BODY.xlsm) (SHA-256: `7B0F6ECEB2DA0C057FA83527F7BC0E6E29773DDA36C93A9B12B3DCB9B6FAEE1D`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY.zip`](./261008_Stock_FG_MA_BODY.zip) (SHA-256: `B5292FB65C091C79980915F5B4AC23F64F3659ADA0739EC4776309AF56B2C4A3`)
* **Bản cập nhật BODYY**: [`261008_Stock_FG_MA_BODYY.xlsm`](./261008_Stock_FG_MA_BODYY.xlsm) (SHA-256: `E4E80F70350257A83FFF50C829D79DB49F9F8F29D299A7B9B42987A67F7DBECD`) | [`ZIP`](./261008_Stock_FG_MA_BODYY.zip) (SHA-256: `42A68E1CE3BFD50A261C8B53B05B5F8635ADCCAF3463D3DC6AB767AD8F4D6824`)
* **Bản phiên hiệu V3.7 đồng bộ**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) | [`ZIP`](./261008_Stock_FG_MA_BODY_V3_7.zip)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7_8.md`](./RELEASE_NOTES_V3_7_8.md)

---

## 🛠 Tính năng nổi bật & Cập nhật mới (V3.7.8 - 21/10/2026)

1. **Khắc phục triệt để lỗi tính toán & hiển thị tỷ lệ % trên Báo Cáo Tổng Thể (`BAO CAO TONG THE`)**:
   - Sửa lỗi điều kiện dừng vòng lặp: Thay thế điều kiện so sánh chuỗi bằng bẫy số nguyên chuẩn xác `IsNumeric(Stt) And Val(Stt) > 0`.
   - Triệt tiêu 100% hiện tượng vòng lặp xóa nhầm dòng `Tổng SL mã`, chèn thêm dòng thừa (dòng 24) và ghi đè giá trị số lên dòng tỷ lệ gây ra lỗi hiển thị `120800%` và `41900%`.
   - Tỷ lệ % tại dòng 23 và dòng 42 hiển thị chuẩn xác từng phân vị (0XE: 28%, <=50XE: 8%, <=100XE: 13%, ..., Tổng: 100%).

2. **Khóa cứng tọa độ Panel Chọn Model & Nút Đổi Shop (Chống xê dịch/drift)**:
   - Cố định kích thước chuẩn các cột M (3), N (22), O (6) và các hàng 1..6 trên sheet `CANH BAO`.
   - Nút **`SHOP: BODY [ BẤM ĐỔI XMĐ ]`** và các nút chức năng được gán thuộc tính `Placement = xlMoveAndSize`, khóa chặt vừa khít bên trong ô `N5:O6`, không còn hiện tượng xê dịch, tràn sang cột M hay che lấp các dòng mô tả và ô chọn checkbox khi nạp lại dữ liệu.
   - Căn giữa toàn bộ Checkbox trong cột O thẳng hàng, thẩm mỹ cao.

3. **Cập nhật danh sách mã vật tư FG mới nhất từ `261021_Ma FG.xlsx`**:
   - Tách riêng 23 mã dùng chung cho 2 dòng xe VF6 và VF7 (mỗi model 1 dòng riêng với `US = 1`) đúng theo chỉ đạo phân tách để tiện theo dõi.
   - Cập nhật 2 mã thay thế EC: `BIW71005061HB` và `BIW71006061HB` (thay thế mã cũ đuôi số trơn).
   - Bổ sung 5 mã mới VF6 (`BIW74239330AA` .. `440AA`), 2 mã VF8 Dập (`BIW30008492`, `30009561`), 1 mã VF9NP (`BIW35111320AA`), 1 mã Limo revision AC (`BIW32011830AC`).
   - Cập nhật US cho `BIW32024120AA` (Limo) = 1, `BIW30210270` (VF6) = 2.
   - Bổ sung cột phân loại **`Khu vực`** (`FG Hàn` / `FG Dập`) tại cột H sheet `DATA`.
   - Bảo toàn 100% dữ liệu Xưởng XMĐ, VF9, Lạc Hồng và các bản RHD.

4. **Tích hợp Webhook riêng cho Báo Cáo Tổng Thể Kaizen (`CAU HINH`)**:
   - Thêm cấu hình **`TEAMS_WEBHOOK_KAIZEN_URL`** tại ô A32 sheet `CAU HINH`.
   - Cơ chế linh hoạt: Nếu người dùng điền URL webhook riêng -> tự động gửi Báo cáo Tổng thể vào nhóm Teams đó; nếu để trống -> tự động sử dụng link Webhook cũ (`TEAMS_WEBHOOK_URL`).

5. **Kích hoạt tự động gửi cảnh báo Teams ngay khi tải tồn xong (Khắc phục lỗi 'tải xong không thấy gửi')**:
   - Tích hợp lệnh gọi gửi cảnh báo Teams (`SendTeamsAlert`) trực tiếp vào cuối quy trình `RefreshStock` ngay sau khi SAP MB52 tải về và tính toán xong (áp dụng cho cả khi bấm nút tải tay lẫn chạy hẹn giờ tự động).
   - Bổ sung cơ chế fallback tự động đếm toàn bộ mã cảnh báo trên sheet khi người dùng chưa tick chọn model cụ thể, không còn bị tình trạng đếm 0 mã và bỏ qua gửi.
