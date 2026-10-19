# SAP VBA Tool — Bộ Công Cụ Excel Tự Động Hóa Sản Xuất & Ra Đơn

Thư mục này tập hợp các tệp công cụ Excel VBA chính thức, mới nhất phục vụ vận hành:

1. **261010_OB_TOOL_V1_RC65.xlsm**: 
   - Công cụ **Ra đơn OB / DO tự động trên SAP** (Bản phát hành RC65 mới nhất).
   - Đã xử lý triệt để: Tự quét VL10D cho nhánh 3001, tự hủy dòng mã ngoài trong OB 3001 trước khi Save (khử lỗi WMS VL589), và vá lỗi sập đọc lại chứng từ VL03N.

2. **260810_RA_DON_NVL_V31_ZONE.xlsm**:
   - Công cụ **Ra đơn Nguyên vật liệu theo Zone** (Bản V31 Zone Line Limo mới nhất).
   - Tối ưu động theo chỉ đạo: Không giới hạn danh mục mã cứng, cho phép mọi mã linh kiện chạy phân bổ FIFO bình thường.
   - Cơ chế gom đơn linh hoạt theo Zone line: **Zone 12 thì gom 12, Zone 22 thì tự gom 22**, tuyệt đối không lẫn lộn Zone trong cùng đơn OB.
   - Cấu trúc thứ bậc gom đơn chuẩn hóa: **ZONE** (Ưu tiên 1) $\rightarrow$ **TÊN TL** (Ưu tiên 2) $\rightarrow$ **TRẠM** (Ưu tiên 3).
   - VBE Compile Native PASS 100%, MIGO 311 precheck xác thực 0 lỗi.

3. **261008_Stock_FG_MA_BODY.xlsm**:
   - Công cụ **Quản trị & Cảnh báo tồn kho Thành phẩm** (Bản chuẩn V3.7).
   - Đã chuẩn hóa: Đổi thứ tự cột (Line Body -> BD Kho), định dạng số nguyên US, font Times New Roman và đổ màu phân biệt text khi copy sang Teams.

---
Xác minh tính toàn vẹn bằng tệp SHA256SUMS.txt.
