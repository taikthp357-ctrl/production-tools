# BẢN PHÁT HÀNH V3.7.12 - TỐI ƯU GIAO DIỆN, CHUẨN HÓA DỮ LIỆU & LỌC BÁO CÁO KAIZEN

**Ngày phát hành**: 21/10/2026  
**Trạng thái**: Production Ready - VBE Compile PASS 100%  
**Tệp chính duy nhất**: `261008_Stock_FG_MA_BODY.xlsm` (không cần duy trì tệp phụ `BODYY`)

---

## 1. CÁC NÂNG CẤP & KHẮC PHỤC TRIỆT ĐỂ (V3.7.12)

### A. Khắc phục lỗi lệch nút và tràn chữ khi co giãn (Panel Chọn Model - Ảnh 1)
- **Vấn đề cũ**: Ô `N3:O4` bị merge và chứa đoạn text dài *"ĐANG XEM: SHOP BODY (Ô TÔ) - Bấm nút dưới để chuyển XMĐ"*. Khi thay đổi kích thước cột/dòng hoặc tỷ lệ phóng to thu nhỏ (zoom), nút bấm `Btn_SwitchShop` (nằm ở N5:O6) bị trôi và đè lên dòng chữ này, gây rối mắt và che khuất thao tác.
- **Khắc phục V3.7.12**:
  * Xóa bỏ hoàn toàn dòng text thừa trong `N3:O4`.
  * Đặt nút bấm chuyển shop `Btn_SwitchShop` vừa vặn hoàn hảo tại `N3:O4` với thuộc tính cố định `Placement = xlMoveAndSize`. Nút tự hiển thị rõ trạng thái shop hiện tại và lời nhắc chuyển shop (*"SHOP: BODY [ BẤM ĐỔI XMĐ ]"*).
  * Dòng 5 trở thành Header chuẩn: Ô `N5` ghi rõ **"Model"** (căn trái), ô `O5` ghi **"Chọn"** (căn giữa) với định dạng nền thanh lịch `RGB(235, 240, 245)`.
  * Danh sách checkbox và tên model bắt đầu từ dòng 6 trở đi, căn chỉnh chính xác tâm ô, không bao giờ bị lệch hay chồng lấn khi co giãn màn hình.

### B. Khắc phục cập nhật ma trận theo Model chọn trên sheet `BAO CAO TONG THE`
- **Vấn đề cũ**: Khi người dùng chỉ tick chọn model (ví dụ chỉ chọn `Limo` và `VF8NP`), bấm nút *"CẬP NHẬT MA TRẬN"* thì bảng tính trên Excel vẫn hiện đủ 13 dòng và tổng số lượng không thay đổi.
- **Khắc phục V3.7.12**:
  * Cập nhật macro `UpdateKaizenTableValues`:
    - Khi có model được chọn: Tự động ẩn các dòng model không được chọn (`Rows(r).Hidden = True`), chỉ hiển thị các dòng xe được tick chọn (`Rows(r).Hidden = False`).
    - Hàng `Tổng SL mã`, hàng phân cụm `Cluster <=200XE / >200XE` và hàng `Tỷ lệ %` được tính toán chính xác 100% chỉ dựa trên các dòng model đang hiển thị.
    - Khi người dùng bỏ chọn hoặc không lọc model: Mở lại toàn bộ 13 dòng xe và tổng thể như bình thường.
  * Áp dụng đồng bộ cho cả Bảng 1 (KHO BODY) và Bảng 2 ((LINE + KHO) BODY).

### C. Chuẩn hóa sheet DATA & Chấm dứt hiện tượng đứng hình/tự reload khi gõ phím
- **Vấn đề cũ**:
  1. Mỗi khi gõ hoặc sửa bất kỳ ô nào trên sheet `DATA`, sự kiện `Workbook_SheetChange` lập tức gọi `RebuildViews`, làm Excel bị đơ 2-3s và tải lại màn hình liên tục.
  2. Cột `Khu vực` bị thiếu (541 dòng trống `None`). Giữa dòng 1762 và dòng 4995 có khoảng trống 3,232 dòng trắng, 12 mã mới bị rớt xuống tận dòng 4995-5006.
- **Khắc phục V3.7.12**:
  * **Chấm dứt reload khi sửa DATA**: Loại bỏ lệnh gọi `RebuildViews` trong sự kiện `Workbook_SheetChange` khi sửa sheet `DATA`. Người dùng nhập liệu mượt mà, tức thì, không bị giật lag.
  * **Dọn dẹp & sắp xếp dữ liệu**:
    - Di chuyển toàn bộ 12 mã mới từ dòng 4995-5006 lên nối tiếp dòng 1763-1774.
    - Xóa sạch 3,232 dòng trống, đưa tổng số dòng dữ liệu về chính xác 1,773 dòng liên tục (dòng 2 đến 1774).
    - Đánh lại số thứ tự STT cột A liên tục từ 1 đến 1773.
    - Điền đầy đủ 100% cột `Khu vực` (Cột H): Trạm `H` hoặc có chữ `Dập` là `FG Dập` (192 mã); tra cứu theo tệp đối soát mới `261021_Ma FG.xlsx`; các trạm hàn/gắn khác là `FG Hàn` (1,581 mã).
    - Cập nhật `FormatStockViews` áp dụng đầy đủ kẻ khung, căn giữa và định dạng Times New Roman cho cả cột 8 (`Khu vực`).

### D. Tối ưu hiển thị cột Cảnh báo - Chấm dứt lỗi đè chữ kép (Ảnh 2)
- **Vấn đề cũ**: Khi có tồn hàng chờ xuất (`pending > 0`), mã nguồn cũ nối chuỗi bằng ký tự ngắt dòng `vbLf` (*"Ưu tiên xuất Body\nHàng chờ xuất sang Body"*). Do chiều cao dòng chỉ có 23pt nên dòng 1 và dòng 2 đè lên nhau, tạo ra hiệu ứng chữ bóng ma (ghosting / double text).
- **Khắc phục V3.7.12**:
  * Tối ưu hiển thị thành 1 dòng duy nhất, rõ ràng và kèm theo chính xác số lượng tồn:
    - Nếu có tồn kho MA: `Ưu tiên xuất Body (Chờ xuất: 100)`
    - Nếu cần kéo line: `Ưu tiên kéo line (Chờ xuất: 120)`
    - Nếu chỉ có hàng chờ xuất đủ đáp ứng: `Hàng chờ xuất sang Body (200)`
    - Nếu MA hết tồn: `MA hết tồn`
  * Mở rộng độ rộng Cột 11 (K) lên 34pt, giúp câu chữ hiển thị thoáng đãng, tuyệt đối không bị đè chữ hay che khuất.

### E. Webhook Kaizen linh hoạt (Fallback tự động)
- Bổ sung cơ chế thông minh: Nếu cấu hình `TEAMS_WEBHOOK_KAIZEN_URL` riêng, hệ thống sẽ gửi báo cáo tổng thể vào kênh đó. Nếu để trống, hệ thống sẽ **tự động dùng Webhook nhóm chung (`TEAMS_WEBHOOK_URL`)**, không còn báo lỗi trùng hay chặn gửi.

---

## 2. BẢNG MÃ BĂM VÀ TỆP TIN XÁC THỰC (SHA-256)

| Tệp tin | Kích thước | SHA-256 Checksum |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | ~1.6 MB | `54ad68387fe8c7570154bdb591c931c1e720d42d19a6d20592ef5038cce04725` |
| `261008_Stock_FG_MA_BODY.zip` | ~900 KB | `831ed1ce2baed067516873878b5e751e477238bfb95641a708f1e4b7b6cdef44` |
