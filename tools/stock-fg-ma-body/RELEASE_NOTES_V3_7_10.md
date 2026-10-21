# BẢN PHÁT HÀNH V3.7.10 - CƠ CHẾ CIRCUIT BREAKER & CHỐNG GỬI LẶP TEAMS WEBHOOK

**Ngày phát hành**: 21/10/2026  
**Trạng thái**: Production Ready - VBE Compile PASS 100%  
**Bảo mật & Ổn định**: Triệt tiêu hiện tượng gửi lặp (duplicate triggers), trang bị Circuit Breaker Cooldown khi lỗi Webhook, chống gửi đè schema Kaizen vào Flow Cảnh báo.

---

## 1. NGUYÊN NHÂN SỰ CỐ GỬI LẶP VÀ THỬ LẠI NHIỀU LẦN

1. **Double Triggers trong AutoRefresh**:
   - Trong phiên bản trước, `AutoRefresh` gọi `RefreshStock` (bên trong `RefreshStock` đã có lệnh gửi webhook). Sau khi `RefreshStock` chạy xong và trả về `AutoRefresh`, macro lại tiếp tục gọi `SendTeamsAlert` và `SendTeamsKaizenReport` một lần nữa, khiến cùng một lần chạy bị bắn 2 request liên tiếp cách nhau vài mili-giây.
2. **Xung đột Webhook khi chưa tách link Kaizen**:
   - Khi cấu hình `TEAMS_WEBHOOK_KAIZEN_URL` để trống, macro cũ tự động fallback về `TEAMS_WEBHOOK_URL` của luồng Cảnh báo. Hai payload có cấu trúc dữ liệu khác nhau (`CANH_BAO_TON_KHO` vs `BODY_KAIZEN_2H`) khiến bước `Parse JSON` trên Power Automate của Flow Cảnh báo bị lỗi không nhận dạng được schema.
3. **Cơ chế Retry mặc định của Power Automate**:
   - Power Automate mặc định bật chính sách thử lại (Retry Policy: Default) thử lại 4 lần liên tiếp (ngay lập tức, 20s, 1m, 3m) khi gặp lỗi HTTP hoặc action thất bại.
4. **Thiếu Cooldown khi Webhook bị lỗi kết nối**:
   - VBA chưa có cơ chế ngắt mạch (Circuit Breaker), nếu Webhook gặp sự cố mạng hoặc lỗi phía máy chủ, các chu kỳ kế tiếp vẫn liên tục gửi dồn dập.

---

## 2. CÁC BIỆN PHÁP KHẮC PHỤC TRIỆT ĐỂ (V3.7.10)

### A. Triệt tiêu kích hoạt kép (Single Invocation Pattern)
- Xóa bỏ hoàn toàn lệnh gọi gửi trùng lặp trong `AutoRefresh`.
- Chu trình tải tồn `RefreshStock` đóng vai trò duy nhất xử lý: Tải MB52 -> Cập nhật bảng -> Gửi cảnh báo ngầm (`SendTeamsAlert True`) -> Lên lịch giờ kế tiếp.

### B. Tách biệt Webhook an toàn (Webhook Isolation)
- Báo cáo Tổng thể Kaizen 2H (`SendTeamsKaizenReport`) **chỉ** tự động gửi khi ô A32 sheet `CAU HINH` (`TEAMS_WEBHOOK_KAIZEN_URL`) có link Webhook riêng và khác với `TEAMS_WEBHOOK_URL`.
- Nếu để trống hoặc trùng với link Cảnh báo: Hệ thống tự động bỏ qua (`SKIPPED`) và ghi nhận vào sheet `LOG`, tuyệt đối không gửi đè gây lỗi schema Flow.

### C. Cơ chế ngắt mạch và chờ sau lỗi (Circuit Breaker & 15-Minute Cooldown)
- Bổ sung biến trạng thái cấp module: `mLastAlertFailTime`, `mLastKaizenFailTime`, `mConsecutiveAlertFails`.
- Khi Webhook gặp sự cố (mã HTTP >= 400 hoặc lỗi mạng): Hệ thống kích hoạt thời gian chờ **15 phút**. Trong 15 phút này, các lượt chạy tự động sẽ tự động bỏ qua (`SKIPPED`) để tránh làm nghẽn/spam mạng.
- Khi người dùng chủ động bấm nút tay ("GỬI CẢNH BÁO TEAMS"), hệ thống cho phép bỏ qua cooldown để kiểm tra kết nối ngay lập tức.

### D. Chống gửi trùng lặp nội dung khi tồn kho không đổi (Duplicate Content Suppression)
- Tự động sinh chuỗi nhận dạng nội dung (Content Signature): Model + Số lượng mã cảnh báo + Danh sách top mã và trạng thái cảnh báo.
- Nếu chu kỳ tự động kế tiếp có nội dung cảnh báo y hệt lần gửi thành công trước đó (trong vòng 120 phút), hệ thống tự động bỏ qua việc gửi tin nhắn lặp lại lên Teams, giữ kênh Teams gọn gàng và tránh nhiễu thông tin.

### E. Hướng dẫn cấu hình Power Automate Flow
- Trên Power Automate Web, mở Flow xử lý Webhook:
  1. Nhấn vào dấu `...` tại action gửi Teams (hoặc action HTTP).
  2. Chọn **Settings** (Cài đặt).
  3. Tại mục **Retry Policy** (Chính sách thử lại), chuyển từ **Default** sang **None** (Không thử lại).
  4. Nhấn **Done** và lưu Flow.

---

## 3. BẢNG MÃ BĂM VÀ TỆP TIN XÁC THỰC (SHA-256)

| Tệp tin | Kích thước | SHA-256 Checksum |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | ~1.6 MB | `33E0F7EC8F2E91C79B33686ECA3BB6B4EC3097D60ABCAF9A4A16BDA617257585` |
| `261008_Stock_FG_MA_BODY.zip` | 906,995 B | `04379AC142A9F21759CFFB39766A8907E7187EB1B592ECC24323E1185B6F9F15` |
| `261008_Stock_FG_MA_BODYY.xlsm` | ~1.5 MB | `2EF86E612F8587C36571071F35171A43A96FD59305567F1F8453AAA9718B5233` |
| `261008_Stock_FG_MA_BODYY.zip` | 862,026 B | `C0E0329CC5A495E09282883AC562782A97D58F050457FDB23F6913FFDD0275CF` |
