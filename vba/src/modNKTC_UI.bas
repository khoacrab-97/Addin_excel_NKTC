Attribute VB_Name = "modNKTC_UI"
'==============================================================================
' NHẬT KÝ TC – hộp thông báo hiện được tiếng Việt.
' MsgBox của VBA đổi chữ sang bảng mã ANSI nên mất dấu ("Nh?t ký") -> gọi thẳng MessageBoxW (Unicode) của Windows.
' Không dùng MsgBox / InputBox của VBA trong add-in (vba/tools/test_vba_ascii.py kiểm tra).
'==============================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

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
