# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7.12 (21/10/2026 - Tối ưu giao diện, Chuẩn hóa dữ liệu & Lọc Báo cáo Kaizen)

* **Tệp Excel chính thức**: [`261008_Stock_FG_MA_BODY.xlsm`](./261008_Stock_FG_MA_BODY.xlsm) (SHA-256: `54ad68387fe8c7570154bdb591c931c1e720d42d19a6d20592ef5038cce04725`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY.zip`](./261008_Stock_FG_MA_BODY.zip) (SHA-256: `831ed1ce2baed067516873878b5e751e477238bfb95641a708f1e4b7b6cdef44`)
* **Bản phiên hiệu V3.7 đồng bộ**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) | [`ZIP`](./261008_Stock_FG_MA_BODY_V3_7.zip)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7_12.md`](./RELEASE_NOTES_V3_7_12.md)

---

## 🛠 Tính năng nổi bật & Cập nhật mới (V3.7.12 - 21/10/2026)

1. **Khắc phục lỗi lệch nút và tràn chữ khi co giãn (Panel Chọn Model - Ảnh 1)**:
   - Xóa bỏ dòng text thừa ở `N3:O4`.
   - Nút `Btn_SwitchShop` đặt vừa khít tại `N3:O4` (`Placement = xlMoveAndSize`), không bị xê dịch khi co giãn hàng/cột hay zoom.
   - Dòng 5 làm Header danh sách model: Ô `N5` ghi **"Model"** (căn trái), ô `O5` ghi **"Chọn"** (căn giữa).
   - Danh sách checkbox model hiển thị từ dòng 6 trở đi với RowHeight 22pt chuẩn xác.

2. **Khắc phục cập nhật ma trận theo Model chọn trên sheet `BAO CAO TONG THE`**:
   - Khi tick chọn model cụ thể (ví dụ: chỉ chọn `Limo` và `VF8NP`), bấm *"CẬP NHẬT MA TRẬN"* sẽ tự động ẩn các dòng model không chọn (`Rows.Hidden = True`), chỉ hiển thị các dòng xe được chọn.
   - Hàng `Tổng SL mã`, hàng phân cụm `Cluster <=200XE / >200XE` và hàng `Tỷ lệ %` được tính toán lại chính xác 100% chỉ dựa trên các dòng model đang hiển thị.
   - Khi không chọn model nào thì mở lại toàn bộ 13 dòng xe.

3. **Chuẩn hóa sheet DATA & Chấm dứt reload khi gõ phím**:
   - Loại bỏ lệnh gọi `RebuildViews` trong sự kiện `Workbook_SheetChange` khi sửa sheet `DATA`. Nhập liệu mã mới mượt mà, không bị đơ giật.
   - Dọn sạch 3,232 dòng trống, di chuyển 12 mã mới lên dòng 1763-1774 nối tiếp liên tục (tổng 1,773 dòng).
   - Đánh lại số thứ tự STT cột A liên tục từ 1 đến 1773.
   - Điền đầy đủ 100% cột `Khu vực` (Cột H): Trạm `H` hoặc có chữ `Dập` là `FG Dập` (192 mã); tra cứu theo tệp đối soát mới `261021_Ma FG.xlsx`; các trạm hàn/gắn khác là `FG Hàn` (1,581 mã).

4. **Tối ưu hiển thị cột Cảnh báo - Chấm dứt lỗi đè chữ kép (Ảnh 2)**:
   - Hiển thị 1 dòng duy nhất, rõ ràng và kèm theo chính xác số lượng tồn (ví dụ: `Ưu tiên xuất Body (Chờ xuất: 100)`).
   - Tăng độ rộng Cột 11 (K) lên 34pt, không còn hiện tượng đè chữ hay lỗi chữ bóng ma.

5. **Webhook Kaizen linh hoạt (Fallback tự động)**:
   - Nếu để trống `TEAMS_WEBHOOK_KAIZEN_URL`, hệ thống tự động gửi Báo cáo Tổng thể vào Webhook nhóm chung (`TEAMS_WEBHOOK_URL`).
