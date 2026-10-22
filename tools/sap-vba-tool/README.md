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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.17 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.17**:
  - **Khắc phục triệt để lỗi nhánh báo cáo 2h/lần gửi sai/thừa model**:
    + Bảng HTML gửi lên Teams khớp 100% với giao diện Excel trên sheet `BAO CAO TONG THE`, loại trừ hoàn toàn các dòng model bị ẩn (`ws.Rows(r).Hidden`).
    + Khử bỏ thuật toán tìm kiếm chuỗi con `InStr` làm lẫn lộn các model có tên chứa nhau (chọn `Limo` không còn bị lẫn `Limo RHD`, chọn `VF6` không còn bị lẫn `VF6 CIS/EU` và `VF6 RHD`).
    + Dòng `Tổng SL mã`, phân cụm `<=200XE` / `>200XE` và dòng Tỷ lệ % được tính toán đồng bộ theo đúng các model đang hiển thị, định dạng tỷ lệ chuẩn 1 số thập phân (e.g. `34.3%`).
    + Nút bấm `BÁO CÁO 2H/LẦN` tự động tính toán lại bảng trước khi gửi.
    + Cơ chế tự động gửi báo cáo Kaizen 2h khi bật chế độ Auto (`AutoRun = True`) vào các mốc giờ chẵn (8h, 10h, 12h, 14h, 16h...) song song với cảnh báo tồn kho.
  - **Cập nhật danh mục mã mới cho BODY**: Đã cập nhật đầy đủ danh sách mã mới cho xưởng BODY (1,639 mã), trong đó nâng cấp 4 mã mới sang revision HB (`BIW73010549HB` và `BIW73010520HB` cho VF5 RHD và VF5).
  - **Phân tách Plant tải tồn MB52 theo Shop (BODY / XMĐ)**: BODY tải đúng 3093 và 3001; XMĐ tải đủ 3093, 2001, 3001.
  - **Chốt chặn bảo vệ độ tươi dữ liệu (Data Freshness Gate 3 Lớp)**: Triệt tiêu vĩnh viễn lỗi gửi tồn cũ lúc 8h khi vừa bật máy. Nếu xuất MB52 lỗi hoặc dữ liệu cũ > 15 phút, chặn gửi Teams 100%.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
