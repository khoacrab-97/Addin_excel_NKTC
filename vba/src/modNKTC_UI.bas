Attribute VB_Name = "modNKTC_UI"
'==============================================================================
' NHẬT KÝ TC – hộp thông báo và tiêu đề hộp thoại hiện được tiếng Việt.
' MsgBox và tiêu đề UserForm của VBA đổi chữ sang bảng mã ANSI nên mất dấu ("Nh?t ký") -> gọi thẳng MessageBoxW (Unicode) của Windows.
' Không dùng MsgBox / InputBox của VBA trong add-in (vba/tools/test_vba_ascii.py kiểm tra).
'==============================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, _
    ByRef hWnd As LongPtr) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, _
    ByRef hWnd As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const WM_SETTEXT As Long = &HC

Private Const MSG_MAX_LENGTH As Long = 3000
Private Const MB_SETFOREGROUND As Long = &H10000

' Như MsgBox (cùng hằng số nút/biểu tượng và giá trị trả về) nhưng hiện đúng tiếng Việt.
Public Function ShowMessage(ByVal text As String, Optional ByVal buttons As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    Dim title As String
    title = "Nhật ký TC"
    If Len(text) > MSG_MAX_LENGTH Then text = Left$(text, MSG_MAX_LENGTH - 1) & ChrW(8230)
    ShowMessage = MessageBoxW(Application.hWnd, StrPtr(text), StrPtr(title), buttons Or MB_SETFOREGROUND)
End Function

' Hộp thông báo chung của add-in (chỉ nút OK).
Public Sub ShowInfo(ByVal text As String, Optional ByVal kind As VbMsgBoxStyle = vbInformation)
    ShowMessage text, kind
End Sub

' Tiêu đề UserForm hiện đúng tiếng Việt: cửa sổ form của VBA là cửa sổ ANSI nên tiêu đề mất dấu ("Xu?t vào form");
' ghi lại tiêu đề bằng DefWindowProcW(WM_SETTEXT) (giữ nguyên Unicode). Gọi trong UserForm_Activate.
Public Sub SetUnicodeTitle(ByVal frm As Object)
#If VBA7 Then
    Dim h As LongPtr
#Else
    Dim h As Long
#End If
    Dim title As String
    On Error Resume Next
    title = frm.Caption
    IUnknown_GetWindow frm, h
    If h = 0 Then h = FindWindowA("ThunderDFrame", title)
    If h <> 0 Then DefWindowProcW h, WM_SETTEXT, 0, StrPtr(title)
End Sub
