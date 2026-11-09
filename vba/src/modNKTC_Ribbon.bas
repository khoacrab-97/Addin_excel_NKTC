Attribute VB_Name = "modNKTC_Ribbon"
'==============================================================================
' NHẬT KÝ TC – tab Ribbon "Nhật ký TC": nhận lệnh từ các nút (vba/ribbon/customUI14.xml).
' Đã chuyển: Tạo bảng, Xuất theo ngày, Phiên bản (kèm kết quả kiểm thử phần lõi); các nút khác đang chuyển sang VBA.
'==============================================================================
Option Explicit

Public Const APP_VERSION As String = "3.0.0-dev"

Private mRibbon As Object

Public Sub RB_OnLoad(ByVal ribbon As Object)
    Set mRibbon = ribbon
End Sub

Public Sub RB_Action(ByVal control As Object)
    On Error GoTo Loi
    Select Case control.Id
        Case "btnTaoBang": CmdTaoBang
        Case "btnXuatNgay": CmdXuatNgay
        Case "btnPhienBan": ShowAbout
        Case Else
            ShowInfo "Chức năng này đang được chuyển sang bản VBA. Trong lúc chờ, anh/chị dùng tab ""Nhật ký thi công"" (bản cũ)."
    End Select
    Exit Sub
Loi:
    ShowInfo "Lỗi: " & Err.Description, vbExclamation
End Sub

Private Sub ShowAbout()
    ShowInfo "Nhật ký TC – phiên bản " & APP_VERSION & vbLf & _
        "Excel " & Application.Version & vbLf & _
        "Thư mục cài đặt: " & ThisWorkbook.Path & vbLf & vbLf & _
        "Kiểm thử phần lõi: " & Left$(NKTC_RunTests(), 300)
End Sub
