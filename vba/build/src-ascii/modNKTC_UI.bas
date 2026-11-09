Attribute VB_Name = "modNKTC_UI"
'==============================================================================
' NHAT KY TC ? hop thong bao hien duoc tieng Viet.
' MsgBox cua VBA doi chu sang bang ma ANSI nen mat dau ("Nh?t ky") -> goi thang MessageBoxW (Unicode) cua Windows.
' Khong dung MsgBox / InputBox cua VBA trong add-in (vba/tools/test_vba_ascii.py kiem tra).
'==============================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const MSG_MAX_LENGTH As Long = 1500
Private Const MB_SETFOREGROUND As Long = &H10000

' Nhu MsgBox (cung hang so nut/bieu tuong va gia tri tra ve) nhung hien dung tieng Viet.
Public Function ShowMessage(ByVal text As String, Optional ByVal buttons As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    Dim title As String
    title = U("Nh\u1eadt k\u00fd TC")
    If Len(text) > MSG_MAX_LENGTH Then text = Left$(text, MSG_MAX_LENGTH - 1) & ChrW(8230)
    ShowMessage = MessageBoxW(Application.hWnd, StrPtr(text), StrPtr(title), buttons Or MB_SETFOREGROUND)
End Function

' Hop thong bao chung cua add-in (chi nut OK).
Public Sub ShowInfo(ByVal text As String, Optional ByVal kind As VbMsgBoxStyle = vbInformation)
    ShowMessage text, kind
End Sub
