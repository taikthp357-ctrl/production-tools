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
- **Tính năng nâng cấp nổi bật V3.7.20**:
  - **Tải tồn MB52 theo Ngữ cảnh Shop & Model siêu tốc (1 - 3 giây)**:
    + **BODY**: Chỉ truy vấn Plant `3093` (Kho MA) và `3001` (Kho Body), chỉ nạp mã của Model BODY đang chọn.
    + **XMĐ**: Truy vấn đủ 3 Plant `3093`, `2001` và `3001` (bao quát toàn diện hàng GA sau sơn), nạp cả mã trước sơn và sau sơn của Model XMĐ đang chọn.
    + **ALL**: Truy vấn cả 3 Plant cho toàn bộ model được chọn.
    + **Bẫy an toàn**: Bật popup nhắc nhở nếu chưa chọn model nào, ngăn chặn truy vấn SAP vô ích.
  - **Khắc phục triệt để lỗi Nút tích (Checkbox) bị lệch và méo mó**:
    + Khóa cố định chiều cao dòng `RowHeight = 22pt` đồng bộ trước khi vẽ Checkbox.
    + Loại bỏ vĩnh viễn lệnh `Rows.AutoFit` trên sheet `CANH BAO` (nguyên nhân gây lồi lõm chiều cao hàng).
    + Chuyển thuộc tính sang `box.Placement = xlMove`, giữ nguyên hình vuông `16x16` pixel sắc nét, không bị co kéo hay méo mó khi lọc dữ liệu.
  - **Sửa công thức tính tồn XMĐ (Bổ sung hàng sau sơn tại Plant 3001)**:
    + Bổ sung tồn kho của `Mã sau sơn` tại Plant `3001` (SLoc `3001`, `30MF`, `1001`, `1012`, `1D01`) vào công thức.
    + Ghi nhận trọn vẹn hơn **6,500 chi tiết sau sơn của GA**, chấm dứt triệt để tình trạng báo còi cọc ảo cho xe máy điện.
  - **Chốt chặn thông minh tại ô B9**:
    + Tự động nhận diện khi người dùng tích thêm Model mới nằm ngoài phạm vi vừa tải MB52.
    + Ô `B9` lập tức hiển thị màu vàng cam nổi bật: *"Có thay đổi model/shop - Bấm [ TẢI TỒN MB52 ]"*, giúp người dùng kiểm soát chính xác độ mới của số liệu.
  - **Bảo toàn trọn vẹn nền tảng V3.7.19**:
    + Nút chuyển xưởng 3 trạng thái `Btn_SwitchShop` (`ALL` -> `BODY` -> `XMĐ` -> `ALL`).
    + Chuẩn hóa Adaptive Card 1.4, Single-Model alert sub-21KB, bảng HTML 11 cột, báo cáo Kaizen 2h tự động vào các mốc giờ chẵn.

---

### 🛡️ Bảng Mã Băm Toàn Vẹn (SHA-256)
Tra cứu chi tiết tại [`SHA256SUMS.txt`](SHA256SUMS.txt).
