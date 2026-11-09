Attribute VB_Name = "modNKTC_Ribbon"
'==============================================================================
' NHAT KY TC ? tab Ribbon "Nhat ky TC": nhan lenh tu cac nut (vba/ribbon/customUI14.xml).
' Da chuyen moi nut tru Kiem tra cap nhat (lam cung bo nap tu cap nhat).
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
        Case "btnChenForm": CmdChenForm
        Case "btnXuatForm": CmdXuatForm
        Case "btnNgayDau": CmdChuyenNgay CMD_FIRST
        Case "btnNgayTruoc": CmdChuyenNgay CMD_PREV
        Case "btnNgaySau": CmdChuyenNgay CMD_NEXT
        Case "btnNgayCuoi": CmdChuyenNgay CMD_LAST
        Case "btnHuongDan": ShowHelp
        Case "btnPhienBan": ShowAbout
        Case Else
            ShowInfo U("Ch\u1ee9c n\u0103ng n\u00e0y \u0111ang \u0111\u01b0\u1ee3c chuy\u1ec3n sang b\u1ea3n VBA. Trong l\u00fac ch\u1edd, anh/ch\u1ecb d\u00f9ng tab ""Nh\u1eadt k\u00fd thi c\u00f4ng"" (b\u1ea3n c\u0169).")
    End Select
    Exit Sub
Loi:
    ShowInfo U("L\u1ed7i: ") & Err.Description, vbExclamation
End Sub

Private Sub ShowAbout()
    ShowInfo U("Nh\u1eadt k\u00fd TC \u2013 phi\u00ean b\u1ea3n ") & APP_VERSION & vbLf & _
        "Excel " & Application.Version & vbLf & _
        U("Th\u01b0 m\u1ee5c c\u00e0i \u0111\u1eb7t: ") & ThisWorkbook.Path & vbLf & vbLf & _
        U("Ki\u1ec3m th\u1eed ph\u1ea7n l\u00f5i: ") & Left$(NKTC_RunTests(), 300)
End Sub
