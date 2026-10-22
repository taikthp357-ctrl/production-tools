# Production & Logistics Automation Tools Suite

Kho lưu trữ tập trung các công cụ tự động hóa, giám sát và điều hành sản xuất, tồn kho xưởng và tích hợp SAP ERP.

---

## 📂 Cấu trúc danh mục công cụ (`tools/`)

```text
tools/
├── ob-tool/              # Công cụ tự động hóa tạo & xử lý Outbound Delivery (OB / DO) trên SAP ERP
├── stock-fg-ma-body/     # Công cụ cảnh báo tồn kho thành phẩm cho Xưởng BODY & Xe Máy Điện (XMĐ)
├── ma-control-fg/        # Hệ thống phần mềm điều hành lắp ráp và logistics MA Control FG (Portable & Web)
└── utilities/            # Các công cụ tiện ích phụ trợ (Biên bản bàn giao BBBG, Cân điện...)
```

---

## 🚀 Truy cập nhanh các bản phát hành mới nhất

### 1. 📊 [Stock FG MA BODY](./tools/stock-fg-ma-body/) (Bản phát hành V3.7.19 - 22/10/2026)
*Khôi phục toàn diện tích hợp Xe Máy Điện (XMĐ) song song cùng BODY (MB52 tải đủ 3 Plant 3093/3001/2001 cho 100% mã, chuyển Shop 3 trạng thái ALL/BODY/XMĐ, hiển thị đủ 34 model, tự động nhận diện Shop trong tiêu đề cảnh báo Teams).*
* **Tệp Excel chính thức**: [`tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm`](./tools/sap-vba-tool/261008_Stock_FG_MA_BODY.xlsm) — SHA-256: `1054b39f38f60acb2d1669fc12e29e7589b5adcd976b477cb6db1e7a5e3fc4a5`
* **Gói nén ZIP**: [`tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip`](./tools/sap-vba-tool/261008_Stock_FG_MA_BODY.zip) — SHA-256: `c2e1a920d82e2e764bf275f89fc6ba9913a56211c129e198b1f36361d9530cb9`
* **Bản phát hành V3.7**: [`tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.xlsm`](./tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.xlsm)
* **Ghi chú phát hành**: [`tools/sap-vba-tool/README.md`](./tools/sap-vba-tool/README.md)

---

### 2. ⚡ [OB Tool V1 RC65](./tools/ob-tool/) (Cập nhật 17/10/2026)
*Quét VL10D 1 lần cho nhánh 3001 (bỏ qua batch thiếu nếu hết STO), hủy dòng mã ngoài trong OB 3001 trước Save triệt tiêu lỗi VL589, vá sập đọc lại VL03N.*
* **Tệp Excel**: [`tools/ob-tool/261010_OB_TOOL_V1_RC65.xlsm`](./tools/ob-tool/261010_OB_TOOL_V1_RC65.xlsm) — SHA-256: `FF7D14D2572FCABE42A9E3459A48AD29A1658CCC61789770BBE05B975427FA22`
* **Gói bàn giao Handoff**: [`tools/ob-tool/RC65_HANDOFF.zip`](./tools/ob-tool/RC65_HANDOFF.zip) — SHA-256: `8576D07EC21BD560E7F76D2120CE4E5D97FAA52491B11B2BBA9881446C35ED0E`
* **Vá mã VBA & Changelog**: [`tools/ob-tool/CHANGELOG_RC64_TO_RC65.md`](./tools/ob-tool/CHANGELOG_RC64_TO_RC65.md)

---

### 3. 🖥️ [MA Control FG Suite](./tools/ma-control-fg/)
*Nền tảng kiểm soát lắp ráp và dữ liệu điều hành xưởng (Zero-Port Architecture).*
* **Bản Portable v1.1.0**: [`tools/ma-control-fg/MA_CONTROL_FG_PORTABLE_ZERO_PORT_v1.1.0.zip`](./tools/ma-control-fg/MA_CONTROL_FG_PORTABLE_ZERO_PORT_v1.1.0.zip)
* **Gói Handoff**: [`tools/ma-control-fg/MA_CONTROL_FG_V1_0_0_HANDOFF.zip`](./tools/ma-control-fg/MA_CONTROL_FG_V1_0_0_HANDOFF.zip)

---

### 4. 🧰 [Tiện ích phụ trợ](./tools/utilities/)
* **Biên bản bàn giao**: [`tools/utilities/261008_BBBG_FG_V2_2_14.zip`](./tools/utilities/261008_BBBG_FG_V2_2_14.zip)
* **Cân điện**: [`tools/utilities/candienmoi.zip`](./tools/utilities/candienmoi.zip)

---

## 🔒 Quản lý quyền truy cập & Phân quyền tài khoản

Repository hỗ trợ phân quyền chi tiết cho từng tài khoản GitHub:
1. Khi repository ở chế độ **Private**, chỉ những tài khoản được mời trong danh sách **Collaborators** mới có quyền xem và tải xuống các tệp.
2. Để cấp quyền cho một tài khoản mới:
   - Vào **Settings** -> **Collaborators** -> Nhấn **Add people**.
   - Tìm theo username hoặc email GitHub của nhân sự cần cấp quyền.
   - Chọn vai trò **Read** (chỉ xem & tải về) hoặc **Write** (được đẩy code/cập nhật).
