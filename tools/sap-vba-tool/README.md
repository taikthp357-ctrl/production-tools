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

#### 3. Quản Trị & Cảnh Báo Tồn Kho Thành Phẩm (Bản Chuẩn V3.7.20 Mới Nhất)
- **Tải trực tiếp (.zip)**: 👉 **[261008_Stock_FG_MA_BODY.zip](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip)**
- **Tệp gốc (.xlsm)**: [`261008_Stock_FG_MA_BODY.xlsm`](https://raw.githubusercontent.com/taikthp357-ctrl/production-tools/main/tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm)
- **Tính năng nâng cấp & vá lỗi V3.7.20**:
  - **Khắc phục dứt điểm lỗi Chặn Cảnh báo Teams sau khi tải MB52 thành công (Freshness Gate Fix)**:
    + Thay thế hàm `V2Config` đọc chuỗi bằng `GetLastSuccessDate()` xử lý đa tầng: số serial Excel thực (`> 40000`), chuỗi ISO chuẩn và chuỗi ngày chuẩn. Triệt tiêu tận gốc lỗi *"Chưa có mốc thời gian tải MB52 thành công"*.
  - **Hợp nhất hàm gửi HTTP độc lập (`SendTeamsHttpPost`) & Khắc phục Status=0**:
    + Khởi tạo HTTP object độc lập mỗi request (`MSXML2.ServerXMLHTTP.6.0` / fallback), cấu hình timeout chuẩn (5s/5s/15s/30s), bắt trọn vẹn lỗi hệ thống thật (`Err.Number: Err.Description`).
  - **Tách biệt độc lập 2 nhánh báo cáo (Decoupling Architecture)**:
    + Tải tồn thành công -> Cảnh báo tức thời được gửi ngay; Báo cáo Kaizen chỉ gửi khi đến mốc giờ chẵn hoặc bấm nút thủ công. Nhánh này lỗi không ảnh hưởng nhánh kia.
  - **Tải tồn MB52 theo Ngữ cảnh Shop & Model siêu tốc (1 - 3 giây)**:
    + BODY chỉ tải Plant 3093 & 3001 cho model đang chọn; XMĐ tải Plant 3093, 2001 & 3001 (hàng GA sau sơn); ALL tải cả 3 Plant.
  - **Khắc phục triệt để lỗi Nút tích (Checkbox) bị lệch và méo mó**:
    + Chuẩn hóa chiều cao dòng `RowHeight = 22pt`, loại bỏ `Rows.AutoFit`, chuyển `box.Placement = xlMove` (14/14 checkbox vuông vắn 16x16 pixel).
    + Màu chữ tiêu đề N6:P6 ("Model", "Ngưỡng DOH") đổi sang xanh đậm `RGB(23, 58, 89)` rõ ràng, sắc nét.
  - **Sửa công thức tính tồn XMĐ (Bổ sung hàng sau sơn tại Plant 3001)**:
    + Ghi nhận hơn 6,500 chi tiết sau sơn của GA vào công thức tính tồn.


