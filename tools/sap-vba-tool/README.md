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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.15 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.15**:
  - **Chốt chặn bảo vệ độ tươi dữ liệu (Data Freshness Gate 3 Lớp)**: Triệt tiêu vĩnh viễn lỗi gửi tồn cũ của 1 tiếng trước khi tắt máy, máy Sleep hoặc vừa mở máy lúc 8h. Nếu xuất MB52 thất bại hoặc dữ liệu tồn cũ hơn 15 phút, hệ thống tự động CHẶN GỬI TEAMS 100%, bảo toàn tính chính xác tuyệt đối.
  - **Cải tiến Ngưỡng DOH đặt cùng cột tích Model theo từng Model**: Trên sheet `CANH BAO`, khối chọn Model mở rộng 3 cột `N: Model`, `O: Chọn [✓]`, `P: Ngưỡng DOH`. Người dùng có thể xem và nhấp đúp gõ sửa trực tiếp số ngưỡng DOH của từng model ngay tại cột P (tự động đồng bộ sang CAU HINH và tính lại ma trận tức thì mà không cần chuyển sheet).
  - **Nhắc nhở khi User bật máy vào ca**: Khi mở file, nếu tồn kho từ ca trước (>60 phút), hệ thống đặt trạng thái nhắc nhở màu vàng để người dùng nạp tồn mới, sau đó mới phát cảnh báo chuẩn xác.
  - **Bảo toàn 100% tất cả các mã cảnh báo (Đầy đủ 592 mã)**: Tuyệt đối không cắt xén, gửi theo từng Model riêng biệt tag đúng PIC phụ trách model, payload an toàn sub-21KB.
  - **Chuẩn hóa 3 nút trên Báo cáo tổng thể & Công thức Excel sống**: Khởi tạo siêu tốc (<0.15s) bằng công thức sống `=SUBTOTAL(109, ...)`.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
