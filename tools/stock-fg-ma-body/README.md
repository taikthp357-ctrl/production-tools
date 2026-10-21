# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7.8 (21/10/2026 - Bản Fix Kaizen Matrix & Panel Lock)

* **Tệp Excel chính thức**: [`261008_Stock_FG_MA_BODY.xlsm`](./261008_Stock_FG_MA_BODY.xlsm) (SHA-256: `792D3BD843770D301D4D5C38DBFD7FB39524A5B280F3A127CDDE7FD396C8FE3C`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY.zip`](./261008_Stock_FG_MA_BODY.zip) (SHA-256: `21138C47223B02BB3B50DA6B6DC2BA49AF3E16F347A11249127E3595A033FA37`)
* **Bản cập nhật BODYY**: [`261008_Stock_FG_MA_BODYY.xlsm`](./261008_Stock_FG_MA_BODYY.xlsm) (SHA-256: `289DBA3C3B2760C23F2F5A31E7E6BAE8EB3FFC07A8604DB5023F2636C185C07E`) | [`ZIP`](./261008_Stock_FG_MA_BODYY.zip) (SHA-256: `E1E746A23056A8A7AA51B9FA3546C85E213ABF826455BD244838E4BA1C8A4C5E`)
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
