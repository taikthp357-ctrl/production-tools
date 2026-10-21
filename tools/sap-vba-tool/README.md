# SAP VBA Tool — Bộ Công Cụ Excel Tự Động Hóa Sản Xuất & Ra Đơn

Thư mục này tập hợp các tệp công cụ Excel VBA chính thức, mới nhất phục vụ vận hành sản xuất. Để tránh trường hợp trình duyệt chặn tải tệp `.xlsm` hoặc yêu cầu đăng nhập/2FA, hệ thống đã đóng gói sẵn các tệp `.zip` với đường dẫn tải trực tiếp (Direct CDN Download) không cần đăng nhập.

---

### 📦 Tải Nhanh Trọn Bộ (Bấm là Tải Ngay — Không Cần Đăng Nhập)
👉 **[Tải Toàn Bộ Công Cụ Mới Nhất (.ZIP)](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/BO_CONG_CU_SAP_VBA_LATEST.zip)** *(Bao gồm cả 3 công cụ bên dưới)*

---

### 📋 Danh Sách Chi Tiết Từng Công Cụ

#### 1. Ra Đơn Nguyên Vật Liệu Theo Zone (Bản V31 Zone Line Limo)
- **Tải trực tiếp (.zip)**: 👉 **[260810_RA_DON_NVL_V31_ZONE.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/260810_RA_DON_NVL_V31_ZONE.zip)**
- **Tệp gốc (.xlsm)**: [`260810_RA_DON_NVL_V31_ZONE.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/260810_RA_DON_NVL_V31_ZONE.xlsm)
- **Tính năng nổi bật**:
  - Không giới hạn danh mục mã cố định, tiếp nhận mọi mã linh kiện trên sheet `ORDER`.
  - Phân bổ Zone động: **Zone 12 gom riêng 12, Zone 22 gom riêng 22**, tuyệt đối không lẫn lộn Zone.
  - Thứ tự ưu tiên gom đơn chuẩn hóa: **ZONE** (Ưu tiên 1) $\rightarrow$ **TÊN TL** (Ưu tiên 2) $\rightarrow$ **TRẠM** (Ưu tiên 3).
  - VBE Debug Compile PASS 100%, MIGO 311 precheck 0 lỗi.

#### 2. Ra Đơn OB / DO Tự Động Trên SAP (Bản RC65 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261010_OB_TOOL_V1_RC65.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261010_OB_TOOL_V1_RC65.zip)**
- **Tệp gốc (.xlsm)**: [`261010_OB_TOOL_V1_RC65.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261010_OB_TOOL_V1_RC65.xlsm)
- **Tính năng nổi bật**:
  - Xử lý triệt để: Quét nhả tồn VL10D cho nhánh 3001, tự động hủy dòng mã ngoài trong OB 3001 trước khi Save (khử lỗi WMS VL589), và vá lỗi sập đọc lại chứng từ VL03N.

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.14 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.14**:
  - **Bảo toàn 100% tất cả các mã cảnh báo (Đầy đủ 592 mã)**: Tuyệt đối không cắt xén, không bỏ bớt bất kỳ mã nào dưới tồn.
  - **Cơ chế gửi theo từng Model riêng biệt (Single-Model Teams Message)**: Mỗi Model gửi 1 tin nhắn riêng biệt tag đúng PIC phụ trách model. Tự động chia đợt (tối đa 30 mã/tin) đối với các model có số lượng mã lớn (VF8, Lạc Hồng, VF9, VF5 RHD).
  - **Triệt tiêu vĩnh viễn lỗi 413 `RequestEntityTooLarge`**: Khống chế toàn bộ payload từ 4.5 KB đến 21.0 KB (tuyệt đối an toàn dưới 28 KB của Teams), kèm độ trễ 1s giữa các tin nhắn chống Rate Limit.
  - **Chuẩn hóa quy trình trên Báo cáo tổng thể**: Gỡ bỏ nút tải tồn thừa trên sheet `BAO CAO TONG THE`, trả lại đúng 3 nút chuẩn. Nút `[ 🔄 CẬP NHẬT MA TRẬN ]` chỉ đọc dữ liệu có sẵn trong `TON TONG` của lần tải MB52 đó và lọc theo model được chọn, chạy tức thì (< 0.15s), không quét lại MB52.
  - **Dữ liệu sống động**: Toàn bộ dòng tổng và tỷ lệ dùng công thức Excel sống (`=SUBTOTAL(109, ...)`, `=SUM(...)`, `=IF(...)`).
  - **Panel chọn Model trên `CANH BAO`**: Có nút `[ ✓ Chọn hết ]`, `[ ✗ Bỏ hết ]`, Master Checkbox và danh sách model sạch sẽ, không ô trống.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
