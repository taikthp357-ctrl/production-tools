# BẢN PHÁT HÀNH V3.7.11 - LỌC BÁO CÁO THEO MODEL CHỌN & SỬA LỖI FONT HTML

**Ngày phát hành**: 21/10/2026  
**Trạng thái**: Production Ready - VBE Compile PASS 100%  
**Tính năng trọng tâm**: Lọc bảng Báo cáo Tổng thể Kaizen theo đúng các Model được tick chọn, loại bỏ hoàn toàn lỗi font chữ mojibake, tách biệt nút Tải tồn MB52 không tự ý gửi Teams.

---

## 1. CÁC ĐIỂM SỬA ĐỔI VÀ KHẮC PHỤC TRIỆT ĐỂ (V3.7.11)

### A. Sửa triệt để lỗi chữ và định dạng (Font & Date Mojibake Fix)
1. **Lỗi biểu tượng cảm xúc bị biến dạng (`ðŸ˜Š`, `áº¯ ï¸0`, `âce...`)**:
   - Nguyên nhân: Việc để ký tự emoji UTF-8 trực tiếp trong mã nguồn `.bas` khiến Excel khi nạp (import qua code page ANSI CP1252) bị phân rã byte thành các ký tự rác.
   - Khắc phục:
     * Tiêu đề bảng dùng mã HTML entity chuẩn `&#128202;` (biểu đồ cột) hoặc chữ chuẩn tiếng Việt.
     * Hàng tổng hợp mức tồn (Cluster) dùng mã HTML entity chuẩn `&#9888; &lt;= 200XE: ...` (cảnh báo vàng) và `&#9989; &gt; 200XE: ...` (tồn an toàn xanh lá).
     * 100% mã ký tự đều là chuẩn 7-bit ASCII trên tệp nguồn VBA, không bao giờ bị lỗi font hay vỡ ký tự trên Teams hoặc Power Automate.
2. **Lỗi ngày tháng hiển thị số serial `46316`**:
   - Khắc phục: Thay thế việc đọc `.Value2` bằng hàm định dạng chuỗi chuẩn `Format(..., "dd/mm/yyyy")`. Tiêu đề bảng giờ đây hiển thị rõ ràng: `21/10/2026`.

### B. Tách biệt hoàn toàn nút "TẢI TỒN MB52" (Manual Loading Isolation)
- **Trước đây**: Khi người dùng bấm nút "TẢI TỒN MB52", macro `RefreshStock` tự ý gọi cả macro gửi Báo Cáo Tổng Thể `SendTeamsKaizenReport` lên Teams.
- **V3.7.11**:
  * Khi người dùng bấm nút tải tồn thủ công trên màn hình: Hệ thống **chỉ tải dữ liệu SAP MB52 và cập nhật bảng tính**, tuyệt đối **không tự ý gửi bất kỳ tin nhắn nào lên Teams**. Người dùng có thể thoải mái xem lại số liệu trước khi quyết định gửi.
  * Chỉ khi chạy chế độ hẹn giờ tự động mỗi 1 tiếng (`AutoRun = True`): Hệ thống mới tự động gửi thẻ cảnh báo ngầm.
  * Báo cáo Tổng thể Kaizen chỉ gửi khi người dùng chủ động bấm nút "GỬI BÁO CÁO TỔNG THỂ TEAMS" trên sheet `BAO CAO TONG THE`.

### C. Chỉ báo cáo các Model được chọn (Model Selection Filter)
- **Áp dụng cho cả Báo Cáo Tổng Thể (Kaizen) và Thẻ Cảnh Báo (Alert)**:
  * Nếu người dùng tick chọn model cụ thể trên panel (ví dụ: chỉ chọn `VF6`):
    - Cả Bảng 1 (KHO BODY) và Bảng 2 ((LINE + KHO) BODY) **chỉ hiển thị đúng dòng của Model được chọn** (và các biến thể xuất khẩu/tay lái nghịch tương ứng như `VF6`, `VF6 CIS/EU`, `VF6 RHD`).
    - Hàng `Tổng SL mã`, hàng phân nhóm `Cluster` và hàng `Tỷ lệ %` được **tự động tính toán lại chính xác 100%** dựa riêng trên các dòng model được chọn.
    - Tiêu đề bảng tự động bổ sung tên model: `BÁO CÁO TỔNG THỂ TỒN KHO BODY (2H/LẦN) - MODEL: VF6`.
  * Nếu người dùng chưa tick chọn model nào: Bảng sẽ tự động hiển thị đầy đủ tất cả 13 dòng xe như tổng thể.

---

## 2. BẢNG MÃ BĂM VÀ TỆP TIN XÁC THỰC (SHA-256)

| Tệp tin | Kích thước | SHA-256 Checksum |
| :--- | :--- | :--- |
| `261008_Stock_FG_MA_BODY.xlsm` | ~1.6 MB | `7B0F6ECEB2DA0C057FA83527F7BC0E6E29773DDA36C93A9B12B3DCB9B6FAEE1D` |
| `261008_Stock_FG_MA_BODY.zip` | 912,809 B | `B5292FB65C091C79980915F5B4AC23F64F3659ADA0739EC4776309AF56B2C4A3` |
| `261008_Stock_FG_MA_BODYY.xlsm` | ~1.5 MB | `E4E80F70350257A83FFF50C829D79DB49F9F8F29D299A7B9B42987A67F7DBECD` |
| `261008_Stock_FG_MA_BODYY.zip` | 869,026 B | `42A68E1CE3BFD50A261C8B53B05B5F8635ADCCAF3463D3DC6AB767AD8F4D6824` |
