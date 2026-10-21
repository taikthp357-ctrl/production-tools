# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7.8 (21/10/2026)

* **Tệp Excel chính thức**: [`261008_Stock_FG_MA_BODY.xlsm`](./261008_Stock_FG_MA_BODY.xlsm) (SHA-256: `CC72B1B6D3E4AA895543647258724169A1CA438F99FC5578BA3B0D029E4CE880`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY.zip`](./261008_Stock_FG_MA_BODY.zip) (SHA-256: `E198D4CC9A4D23325F5AFE1ECF5835E958BE9B927DAF9621361538572627EF3A`)
* **Bản cập nhật BODYY**: [`261008_Stock_FG_MA_BODYY.xlsm`](./261008_Stock_FG_MA_BODYY.xlsm) (SHA-256: `6D8A72F13D3041AD1E2CD3DE09DFB14B57A2D895B11882A2BBB70525C9A41E6E`) | [`ZIP`](./261008_Stock_FG_MA_BODYY.zip) (SHA-256: `EDAC537375938851AA87DB9B8E694DDF6B2B8F296B9DD014B0BD316A12B8C7E0`)
* **Bản phiên hiệu V3.7 đồng bộ**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) | [`ZIP`](./261008_Stock_FG_MA_BODY_V3_7.zip)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7.md`](./RELEASE_NOTES_V3_7.md)

---

## 🛠 Tính năng nổi bật & Cập nhật mới (V3.7.8 - 21/10/2026)

1. **Sheet Báo Cáo Tổng Thể Kaizen 2H (`BAO CAO TONG THE`)**:
   - Tự động thống kê ma trận phân bổ dải số xe (0, 1-20, 21-50, 51-100, 101-150, 151-200, 201-300, 301-400, >400 xe) theo từng Model xe (VF3, VF5, VF6, VF7, VF8, VF9, VFe34, EC...).
   - Bao gồm 2 bảng ma trận chuẩn:
     * **Bảng 1**: Thống kê theo `KHO BODY`.
     * **Bảng 2**: Thống kê theo `(LINE + KHO) BODY`.
   - Phân nhóm tổng hợp trực quan: Cột `<= 200XE` (nhóm ưu tiên kiểm soát nguy cơ) và Cột `> 200XE` (nhóm an toàn).
   - Tích hợp nút bấm **`LÀM MỚI BÁO CÁO`** và **`GỬI BÁO CÁO TEAMS (2H)`** ngay trên sheet.

2. **Khắc phục triệt để lỗi trôi thời gian hẹn giờ (Anti-Drift Fixed Wall-Clock Timer)**:
   - Cơ chế neo mốc tròn đồng hồ (`GetNextRoundClock`): 
     * Chu kỳ 60 phút neo chính xác: `07:00, 08:00, 09:00, 10:00...`
     * Chu kỳ 120 phút neo chính xác các mốc chẵn: `06:00, 08:00, 10:00, 12:00, 14:00, 16:00, 18:00...`
   - Dù tác vụ chạy mất bao nhiêu giây/phút, thời gian kích hoạt kế tiếp luôn khóa cố định theo mặt đồng hồ thực, không bao giờ bị trôi lệch tích lũy qua các ngày.

3. **Cập nhật nội dung cảnh báo Teams**:
   - Đổi tiêu đề nhóm cảnh báo từ `... MÃ THIẾU TỒN` thành `... mã Ưu tiên Xuất`.
   - Bảng cảnh báo hiển thị đầy đủ 100% tất cả các mã vật tư đạt điều kiện cảnh báo, không cắt bớt để đảm bảo người dùng nắm trọn vẹn thông tin mà không cần mở file tra cứu.
   - Cột **Tồn MA** hiển thị nền trắng chữ rõ ràng (không tô màu nhầm lẫn với hàng thiếu), 3 cột còn lại (`Line Body`, `BD Kho`, `Khả năng lắp ráp`) giữ nguyên màu cảnh báo nổi bật.
   - Tăng kích thước phông chữ và căn lề bảng to, rõ ràng hơn.

4. **Hỗ trợ Universal Theme (Tương thích tuyệt đối cả Dark Mode và Light Mode)**:
   - Bọc toàn bộ bảng HTML gửi qua Webhook vào thẻ chứa nền trắng chuẩn `#FFFFFF` với mã màu chữ cố định.
   - Đảm bảo hiển thị hoàn hảo, không bị vỡ phông, tương phản ngược hay chìm chữ khi người dùng sử dụng Microsoft Teams ở Giao diện Sáng (Light Theme) lẫn Giao diện Tối (Dark Theme).
   - Dung lượng gói HTML siêu nhẹ (~9.1 KB), an toàn 100% dưới ngưỡng 28 KB của Teams Incoming Webhook.

5. **Đồng bộ Master Data 1,761 mã vật tư mới nhất**:
   - Toàn vẹn 100% dữ liệu gốc, hỗ trợ đầy đủ các mã vật tư chuyển đổi mới nhất từ file `261008_Stock_FG_MA_BODYY.xlsm`.
