# Stock FG MA BODY V3.7.8 — Báo Cáo Tổng Thể Kaizen 2H & Anti-Drift Timer

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256 (21/10/2026)

| Tên Tệp | Kích thước | SHA-256 |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | 938,836 bytes | `CC72B1B6D3E4AA895543647258724169A1CA438F99FC5578BA3B0D029E4CE880` |
| `261008_Stock_FG_MA_BODY.zip` | 853,572 bytes | `E198D4CC9A4D23325F5AFE1ECF5835E958BE9B927DAF9621361538572627EF3A` |
| `261008_Stock_FG_MA_BODYY.xlsm` | 938,842 bytes | `6D8A72F13D3041AD1E2CD3DE09DFB14B57A2D895B11882A2BBB70525C9A41E6E` |
| `261008_Stock_FG_MA_BODYY.zip` | 853,565 bytes | `EDAC537375938851AA87DB9B8E694DDF6B2B8F296B9DD014B0BD316A12B8C7E0` |

---

## 2. Các Tính Năng Nổi Bật & Cải Tiến

### A. Sheet Báo Cáo Tổng Thể Kaizen 2H (`BAO CAO TONG THE`)
- Tự động thống kê ma trận phân bổ dải số xe (0, 1-20, 21-50, 51-100, 101-150, 151-200, 201-300, 301-400, >400 xe) theo từng Model (VF3, VF5, VF6, VF7, VF8, VF9, VFe34, EC...).
- Cung cấp 2 bảng ma trận chuẩn:
  * **Bảng 1**: Thống kê theo `KHO BODY`.
  * **Bảng 2**: Thống kê theo `(LINE + KHO) BODY`.
- Phân nhóm tổng hợp: Cột `<= 200XE` (nhóm ưu tiên kiểm soát) và `> 200XE` (nhóm an toàn).
- Nút bấm **`LÀM MỚI BÁO CÁO`** và **`GỬI BÁO CÁO TEAMS (2H)`** tích hợp trực tiếp trên sheet.

### B. Bộ Hẹn Giờ Chống Trôi Giờ (Anti-Drift Fixed Wall-Clock Timer)
- Sử dụng hàm `GetNextRoundClock(intervalMinutes)` để neo thời gian chạy kế tiếp vào các mốc giờ tròn trên mặt đồng hồ thực:
  * Mốc 60 phút: `07:00, 08:00, 09:00, 10:00...`
  * Mốc 120 phút (2 giờ): `06:00, 08:00, 10:00, 12:00, 14:00, 16:00, 18:00...`
- Triệt tiêu hoàn toàn hiện tượng trôi giờ tích lũy (drift time) do thời gian xử lý macro gây ra.

### C. Nâng Cấp Bảng Cảnh Báo Teams Webhook
- Đổi nhãn `... MÃ THIẾU TỒN` thành `... mã Ưu tiên Xuất`.
- Hiển thị đầy đủ 100% tất cả các mã cảnh báo trong bảng Teams, không cắt bớt.
- Cột **Tồn MA** hiển thị nền trắng chữ rõ ràng (không tô màu nhầm lẫn với thiếu hàng); 3 cột `Line Body`, `BD Kho`, `Khả năng lắp ráp` giữ màu sắc cảnh báo trực quan.
- Tăng kích thước phông chữ và căn lề bảng rộng rãi, dễ đọc.

### D. Hỗ Trợ Giao Diện Sáng & Tối (Universal Theme)
- Khóa cứng màu nền `#FFFFFF` và màu chữ cụ thể cho từng thẻ bảng HTML.
- Tương thích tối đa với cả Microsoft Teams Dark Mode và Light Mode mà không bị chìm màu hay lóa mắt.
- Dung lượng payload HTML chỉ ~9.1 KB, tối ưu an toàn tuyệt đối dưới ngưỡng 28 KB của Teams Webhook.

### E. Đồng Bộ Dữ Liệu 1,761 Mã Vật Tư
- Bảo toàn toàn bộ dữ liệu vật tư và các cấu hình mới nhất từ file gốc `261008_Stock_FG_MA_BODYY.xlsm`.
