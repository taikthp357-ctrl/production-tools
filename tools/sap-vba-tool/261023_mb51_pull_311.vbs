'=============================================================================
' TOOL TỰ ĐỘNG TẢI DỮ LIỆU MB51 (MOVEMENT TYPE 311/312) TỪ SAP VÀO EXCEL
' Nhà máy ô tô VinFast Hải Phòng (NMVFHP) - Plant 3001
' Kho xuất: 1003 | Kho nhận (Line): 3003
'=============================================================================

Option Explicit

Dim SapGuiAuto, application, connection, session
Dim WshShell, fso, targetExcelPath, strDate

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Đường dẫn tới file Excel báo cáo
targetExcelPath = "G:\AI-Manager\projects\ma-control\production-tools\tools\sap-vba-tool\261023_Tracking_CapLine_MB51_vs_ThucTe.xlsx"

' Ngày hiện tại định dạng DD.MM.YYYY
strDate = Right("0" & Day(Date), 2) & "." & Right("0" & Month(Date), 2) & "." & Year(Date)

' 1. Kết nối SAP GUI Session
On Error Resume Next
Set SapGuiAuto  = GetObject("SAPGUI")
If Err.Number <> 0 Or Not IsObject(SapGuiAuto) Then
    MsgBox "LỖI: Không tìm thấy SAP Logon đang chạy!" & vbCrLf & "Vui lòng mở SAP Logon và đăng nhập trước khi chạy script.", vbCritical, "SAP Connection Error"
    WScript.Quit
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or Not IsObject(application) Then
    MsgBox "LỖI: Không lấy được SAP Scripting Engine. Vui lòng kiểm tra quyền Scripting trên SAP GUI.", vbCritical, "SAP Scripting Error"
    WScript.Quit
End If

Set connection = application.Children(0)
If Err.Number <> 0 Or Not IsObject(connection) Then
    MsgBox "LỖI: Không tìm thấy kết nối SAP nào đang hoạt động.", vbCritical, "SAP Connection Error"
    WScript.Quit
End If

Set session = connection.Children(0)
If Err.Number <> 0 Or Not IsObject(session) Then
    MsgBox "LỖI: Không tìm thấy cửa sổ SAP Session đang mở.", vbCritical, "SAP Session Error"
    WScript.Quit
End If
On Error GoTo 0

' 2. Mở T-code MB51
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
session.findById("wnd[0]").sendVKey 0

' 3. Thiết lập thông số lọc trên màn hình MB51
' Plant 3001
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "3001"

' Storage location nhận hoặc xuất: 3003
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "3003"

' Movement Type: 311 (hoặc qua nút Multiple Selection để chọn 311 & 312)
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "311"
session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "312"

' Posting Date: Ngày hôm nay
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = strDate
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = strDate

' Chạy thực thi (F8)
session.findById("wnd[0]/tbar[1]/btn[8]").press

' 4. Thông báo hoàn tất
MsgBox "Đã xuất dữ liệu MB51 thành công cho Plant 3001 - Kho 3003 ngày " & strDate & "!" & vbCrLf & _
       "Vui lòng copy kết quả vào Sheet 'RAW_MB51' của file Excel.", vbInformation, "Hoàn tất Tải SAP MB51"
