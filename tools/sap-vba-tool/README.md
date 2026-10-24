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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.21 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nâng cấp nổi bật V3.7.21**:
  - **Nâng cấp thuật toán chia tin động theo dung lượng byte UTF-8 (Dynamic Chunking <= 25 KB)**:
    + Bỏ giới hạn cố định 30 mã/tin. Dựng toàn bộ Adaptive Card 1.4 + Tag PIC + Header + Bảng HTML 11 cột và đo chính xác dung lượng byte UTF-8 qua `ADODB.Stream`.
    + Thiết lập ngưỡng an toàn 25 KB (25,600 bytes) chừa 3 KB đệm so với trần 28 KB của Teams.
    + Nếu dung lượng <= 25 KB, gửi chung **1 tin nhắn duy nhất** (ví dụ Model VF6 có 31 mã ~13.3 KB gửi trọn vẹn trong 1 tin, không còn bị xé lẻ 1 mã sang tin thứ 2).
    + Khi vượt 25 KB (model có >55-60 mã), tự động chia phần cân đối: `VF6 — Phần 1/2`, `VF6 — Phần 2/2` (đảm bảo không trùng, không sót mã).
    + Ghi LOG chi tiết: `[TEAMS_ALERT] PASS | Model VF6 (1/1): 31/31 ma - 13.3 KB (13619 B) - HTTP 202`.
  - **Bảo toàn toàn bộ các tính năng V3.7.20**:
    + Freshness gate `GetLastSuccessDate` đọc ngày serial Excel không bị chặn gửi.
    + Hàm gửi HTTP độc lập `SendTeamsHttpPost` bắt trọn lỗi thật.
    + Tải MB52 theo ngữ cảnh Shop & Model siêu tốc 1-3s (BODY: 3093/3001, XMĐ: 3093/2001/3001 bao quát hàng GA sau sơn).
    + Khắc phục triệt để Checkbox lệch méo (RowHeight 22pt, `box.Placement = xlMove`).
    + Màu chữ tiêu đề Model xanh đậm `RGB(23, 58, 89)`.

#### 4. Đối Soát Cấp Line & Soi Lệch MB51 (Bản Chuẩn V1 Gắn Trực Tiếp Tool - NMVFHP)
- **Tải trực tiếp trọn bộ (.zip)**: 📥 **[261023_Tracking_CapLine_MB51_vs_ThucTe.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261023_Tracking_CapLine_MB51_vs_ThucTe.zip)** *(Chứa file .xlsm có gắn sẵn Macro 1-Click + file .xlsx)*
- **Tệp Excel Macro Gắn Sẵn Tool (.xlsm)**: 📊 **[261023_Tracking_CapLine_MB51_vs_ThucTe.xlsm](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261023_Tracking_CapLine_MB51_vs_ThucTe.xlsm)**
- **Tính năng tích hợp trực tiếp (All-In-One)**:
  - **Nút bấm 1-Click gắn trực tiếp trên Sheet BC_GAP và RAW_MB51**:
    + [ 🔄 TẢI DỮ LIỆU SAP MB51 ]: Tự động kết nối SAP GUI, mở MB51 lọc Plant 3001, Kho nhận 3003, Movement Types 311/312 ngày hôm nay, trích xuất dữ liệu thẳng vào sheet RAW_MB51 và làm mới báo cáo. Không cần chạy script bên ngoài!
    + [ 📊 LÀM MỚI BÁO CÁO ]: Tính toán lại toàn bộ dữ liệu đối soát.
    + [ 🧹 XÓA DỮ LIỆU RAW ]: Dọn sạch bảng dữ liệu cũ trên sheet RAW_MB51.
  - **2 Dạng báo cáo chuyên biệt**:
    + **Báo cáo Chênh lệch GAP (BC_GAP)**: Soi chi tiết theo giờ, 3 thẻ KPI tổng, tự động highlight đỏ/vàng khi phát sinh chênh lệch.
    + **Báo cáo Tiến độ Line Nhận (BC_THUC_NHAN)**: Thống kê sản lượng thực tế nhận được theo từng khung giờ.
  - **Master data chuẩn hóa**: Tích hợp danh mục mã và thông số từ 261015_Ord hang MFG (2).xlsb (Plant 3001, Kho 1003 sang 3003).
