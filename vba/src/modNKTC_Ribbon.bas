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
            MsgBox "Chức năng này đang được chuyển sang bản VBA. Trong lúc chờ, anh/chị dùng tab ""Nhật ký thi công"" (bản cũ).", _
                vbInformation, "Nhật ký TC"
    End Select
    Exit Sub
Loi:
    MsgBox "Lỗi: " & Err.Description, vbExclamation, "Nhật ký TC"
End Sub

Private Sub ShowAbout()
    MsgBox "Nhật ký TC – phiên bản " & APP_VERSION & vbLf & _
        "Excel " & Application.Version & vbLf & _
        "Thư mục cài đặt: " & ThisWorkbook.Path & vbLf & vbLf & _
        "Kiểm thử phần lõi: " & Left$(NKTC_RunTests(), 300), vbInformation, "Nhật ký TC"
End Sub
