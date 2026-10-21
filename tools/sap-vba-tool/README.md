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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.13 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.13**:
  - **Khắc phục triệt để lỗi đường cảnh báo Teams (`RequestEntityTooLarge` 413)**: Khống chế Top 15 mã cảnh báo khẩn cấp nhất, giảm kích thước payload từ 250 KB xuống chỉ còn ~11 KB (an toàn tuyệt đối dưới ngưỡng 28 KB của Power Automate).
  - **Dữ liệu sống cho Báo cáo tổng thể Kaizen**: Chuyển đổi 100% các ô tổng và tỷ lệ sang công thức Excel động (`=SUBTOTAL(109, ...)`, `=SUM(...)`, `=IF(...)`), tự động cập nhật số liệu khi ẩn/hiện model.
  - **Tối ưu Panel chọn Model**: Bổ sung 2 nút bấm thao tác nhanh `[ ✓ Chọn hết ]` và `[ ✗ Bỏ hết ]` cùng **Master Checkbox** tại dòng tiêu đề, xóa bỏ cảm giác trống ô phía trên và cho phép lọc model chỉ bằng 1 cú nhấp chuột.
  - **Nút tải tồn trực tiếp**: Thêm nút `[ 📥 TẢI TỒN MB52 ]` ngay trên sheet `BAO CAO TONG THE`.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
