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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.18 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nổi bật V3.7.18**:
  - **Khắc phục triệt để lỗi khi gửi Cảnh báo Teams (SendTeamsAlert)**:
    + **Bảo toàn cơ chế gửi theo từng Model riêng biệt (Single-Model Alert Dispatch)**: Tối đa 30 mã/tin nhắn giúp khống chế payload an toàn sub-21KB (thực tế ~8.5KB), loại bỏ hoàn toàn nguy cơ quá tải webhook và lỗi HTTP 413 Payload Too Large. Tag đúng PIC phụ trách của từng model.
    + **Chuẩn hóa schema Adaptive Card 1.4**: Hạ chuẩn từ 1.5 về 1.4, loại bỏ thẻ `Table` không tương thích với Teams Power Automate connector, thay thế bằng Card TextBlock tóm tắt gọn gàng, tương thích 100% với mọi nền tảng Teams/Power Automate.
    + **Bảo toàn bảng HTML 11 cột hiển thị hoàn hảo**: Toàn bộ chi tiết mã cảnh báo được nhúng trọn vẹn qua trường `table_html` với định dạng màu sắc chuyên nghiệp.
    + **Bổ sung định danh `report_type: "STOCK_ALERT"`**: Giúp Power Automate phân luồng chính xác giữa Cảnh báo tồn kho (`STOCK_ALERT`) và Báo cáo Kaizen 2h (`BODY_KAIZEN_2H`).
    + **Nâng cấp bộ lọc Model sang từ điển chính xác `SelectedModels`**: Khử bỏ hoàn toàn tìm kiếm chuỗi con `InStr`, chỉ gửi đúng các model được chọn trong ca (2-3 model), không còn tình trạng gửi nhầm/thừa model.
    + **Thêm cơ chế tự động Calculate, giãn cách 1.2s và thử lại (Retry 2 lần)**: Tránh nghẽn mạng và đảm bảo số liệu cảnh báo luôn mới nhất trước khi phát tin.
  - **Khắc phục nhánh báo cáo 2h/lần (V3.7.17)**: Bảng HTML khớp 100% Excel, loại trừ model ẩn, tự động gửi vào các mốc giờ chẵn khi bật Auto.
  - **Cập nhật danh mục mã mới cho BODY**: Đã cập nhật đầy đủ danh sách mã mới cho xưởng BODY (1,639 mã), nâng cấp revision HB (`BIW73010549HB` và `BIW73010520HB`).
  - **Phân tách Plant tải tồn MB52 theo Shop (BODY / XMĐ)**: BODY tải đúng 3093 và 3001; XMĐ tải đủ 3093, 2001, 3001.
  - **Chốt chặn bảo vệ độ tươi dữ liệu (Data Freshness Gate 3 Lớp)**: Triệt tiêu vĩnh viễn lỗi gửi tồn cũ lúc 8h khi vừa bật máy. Nếu xuất MB52 lỗi hoặc dữ liệu cũ > 15 phút, chặn gửi Teams 100%.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
