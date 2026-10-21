# Stock FG MA BODY V3.7.8 — Báo Cáo Tổng Thể Kaizen 2H & Anti-Drift Timer

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256 (Cập nhật 21/10/2026)

| Tên Tệp | Kích thước | SHA-256 |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | 940,328 bytes | `E92DF180998FA41AD23A7FE2A52236B71FED14DDD545022F572AE4DB7C646BC6` |
| `261008_Stock_FG_MA_BODY.zip` | 856,486 bytes | `1F82F42D79EDF39FCF7BE887DC101D43FD194063EE023A6537C8AABC8B5E842E` |
| `261008_Stock_FG_MA_BODYY.xlsm` | 940,334 bytes | `E7E30461EED14901D6BA323F3656DCC0B6B86F822CFC1082BDE27BC1CF1D236D` |
| `261008_Stock_FG_MA_BODYY.zip` | 856,471 bytes | `3E8610CBCAA093E70C3F9C67B332203371413D6C46CFFF87FCC814C5AC3BFBC9` |

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

### E. Kích hoạt tự động gửi cảnh báo Teams ngay khi tải tồn xong
- **Nguyên nhân**: Trong phiên bản trước, quy trình tải tồn kho `RefreshStock` khi hoàn thành xuất/nhập MB52 chỉ gọi tính toán tồn kho và lên lịch hẹn giờ kế tiếp mà không gọi `SendTeamsAlert`. Lệnh gửi cảnh báo chỉ nằm trong vòng lặp tự động `AutoRefresh`. Do đó khi người dùng tải tồn thủ công hoặc khi chưa bật AUTO, hệ thống tải xong nhưng không gửi tin nhắn.
- **Khắc phục**:
  * Đặt lệnh gọi `SendTeamsAlert AutoRun` và `SendTeamsKaizenReport AutoRun` trực tiếp ngay sau khi `Export_MB52` hoàn tất trong `RefreshStock`.
  * Chuẩn hóa logic đếm cảnh báo: Nếu chưa tick chọn model cụ thể, hệ thống tự động fallback sang chế độ `ALL` để gửi trọn vẹn tất cả các mã cảnh báo hiện có trên sheet `CANH BAO`, loại bỏ hoàn toàn lỗi đếm 0 mã và bị skip.
