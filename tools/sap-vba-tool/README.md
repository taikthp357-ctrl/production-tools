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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.19 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.19**:
  - **Khôi phục toàn diện dữ liệu & quy trình Xe Máy Điện (XMĐ) song song cùng BODY**:
    + **Tải tồn MB52 đầy đủ 100% cả 3 Plant (`3093`, `3001`, `2001`)**: Truy vấn đồng thời toàn bộ 1,268 mã chi tiết của cả BODY (1,639 dòng) và XMĐ (134 dòng). Không còn tình trạng MB52 chỉ tải riêng BODY làm mất tồn của XMĐ.
    + **Chế độ xem Shop 3 trạng thái linh hoạt (ALL / BODY / XMĐ)**: Mặc định hiển thị `ALL` (Cả BODY & XMĐ). Nút bấm chuyển Shop `Btn_SwitchShop` tại `N3:P4` hỗ trợ chuyển đổi tuần hoàn 3 chế độ: `TẤT CẢ (BODY & XMĐ)` -> `BODY` -> `XMĐ` -> `TẤT CẢ`.
    + **Đầy đủ 34 Model trên Panel chọn & Ngưỡng DOH**: Khi ở chế độ `ALL`, hiển thị đủ 13 model BODY và 21 model XMĐ cùng ô Checkbox và cột Ngưỡng DOH riêng biệt cho từng model.
    + **Định danh chính xác Shop trong Cảnh báo Teams**: Hàm `GetModelShop` tự động nhận diện đúng Model thuộc xưởng BODY hay XMĐ để hiển thị tiêu đề chuẩn xác (e.g. `CẢNH BÁO TỒN KHO XMĐ - MODEL VF6,7`), tag đúng PIC phụ trách.
  - **Bảo toàn toàn bộ cải tiến V3.7.18**:
    + Chuẩn hóa Adaptive Card 1.4, bảng chi tiết 11 cột hiển thị hoàn hảo qua `table_html`, Single-Model dispatch sub-21KB (thực tế ~8.5KB) chống quá tải webhook.
    + Báo cáo Kaizen 2h khớp 100% Excel, loại trừ model ẩn, tự động gửi vào mốc giờ chẵn.
    + Chốt chặn độ tươi dữ liệu (Data Freshness Gate 3 Lớp), triệt tiêu lỗi gửi tồn cũ lúc 8h sáng.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
