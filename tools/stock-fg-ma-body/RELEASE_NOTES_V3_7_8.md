# Stock FG MA BODY V3.7.8 — Báo Cáo Tổng Thể Kaizen 2H & Anti-Drift Timer

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256 (Cập nhật 21/10/2026)

| Tên Tệp | Kích thước | SHA-256 |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | 940,328 bytes | `792D3BD843770D301D4D5C38DBFD7FB39524A5B280F3A127CDDE7FD396C8FE3C` |
| `261008_Stock_FG_MA_BODY.zip` | 856,486 bytes | `21138C47223B02BB3B50DA6B6DC2BA49AF3E16F347A11249127E3595A033FA37` |
| `261008_Stock_FG_MA_BODYY.xlsm` | 940,334 bytes | `289DBA3C3B2760C23F2F5A31E7E6BAE8EB3FFC07A8604DB5023F2636C185C07E` |
| `261008_Stock_FG_MA_BODYY.zip` | 856,471 bytes | `E1E746A23056A8A7AA51B9FA3546C85E213ABF826455BD244838E4BA1C8A4C5E` |

---

## 2. Các Tính Năng Nổi Bật & Vá Lỗi Chi Tiết

### A. Vá lỗi tính toán & hiển thị tỷ lệ % trên Sheet Báo Cáo Tổng Thể (`BAO CAO TONG THE`)
- **Nguyên nhân gốc rễ**: Vòng lặp `UpdateKaizenTableValues` so sánh chuỗi `"Tổng"` và mã escape `VN()` không khớp chuỗi unicode runtime, dẫn đến việc xem dòng 21 (`Tổng SL mã`) là một dòng model và xóa dữ liệu; sau đó chèn thêm dòng 22 và ghi đè giá trị số lên dòng tỷ lệ vốn có định dạng `0%` (gây lỗi `120800%` và `41900%`).
- **Khắc phục**:
  * Thay thế điều kiện dừng bằng `IsNumeric(Stt) And Val(Stt) > 0`.
  * Khóa cứng vị trí dòng `Tổng SL mã` (dòng 21, 40), dòng cụm (dòng 22, 41) và dòng `Tỷ lệ` (dòng 23, 42).
  * Định dạng số nguyên chuẩn `#,##0` cho dòng tổng kết và cụm; định dạng `"0%"` chuẩn xác cho dòng tỷ lệ (0XE: 28%, <=50XE: 8%, <=100XE: 13%...).

### B. Khóa cứng vị trí Panel Chọn Model & Nút Đổi Shop (Chống trôi dạt)
- Cố định kích thước chuẩn các cột M (3), N (22), O (6) và các hàng 1..6 trên sheet `CANH BAO`.
- Nút **`SHOP: BODY [ BẤM ĐỔI XMĐ ]`** và các nút chức năng được gán `Placement = xlMoveAndSize`, khóa chặt vừa khít trong ô `N5:O6`, không còn hiện tượng xê dịch, tràn sang cột M hay che lấp các dòng mô tả và ô chọn checkbox khi nạp lại dữ liệu.
- Căn giữa toàn bộ Checkbox trong cột O thẳng hàng, thẩm mỹ cao.

### C. Cập nhật danh sách mã vật tư FG mới nhất từ `261021_Ma FG.xlsx`
- Tách riêng 23 mã dùng chung cho 2 dòng xe VF6 và VF7 (mỗi model 1 dòng riêng với `US = 1`) đúng theo chỉ đạo phân tách để tiện theo dõi.
- Cập nhật 2 mã thay thế EC: `BIW71005061HB` và `BIW71006061HB`.
- Bổ sung 5 mã mới VF6 (`BIW74239330AA` .. `440AA`), 2 mã VF8 Dập (`BIW30008492`, `30009561`), 1 mã VF9NP (`BIW35111320AA`), 1 mã Limo revision AC (`BIW32011830AC`).
- Cập nhật US cho `BIW32024120AA` (Limo) = 1, `BIW30210270` (VF6) = 2.
- Bổ sung cột phân loại **`Khu vực`** (`FG Hàn` / `FG Dập`) tại cột H sheet `DATA`.
- Bảo toàn 100% dữ liệu Xưởng XMĐ, VF9, Lạc Hồng và các bản RHD.

### D. Tích hợp Webhook riêng cho Báo Cáo Tổng Thể Kaizen (`CAU HINH`)
- Thêm cấu hình **`TEAMS_WEBHOOK_KAIZEN_URL`** tại ô A32 sheet `CAU HINH`.
- Cơ chế linh hoạt: Nếu người dùng điền URL webhook riêng -> tự động gửi Báo cáo Tổng thể vào nhóm Teams đó; nếu để trống -> tự động sử dụng link Webhook cũ (`TEAMS_WEBHOOK_URL`).
