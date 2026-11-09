Attribute VB_Name = "modNKTC_Commands"
'==============================================================================
' NHAT KY TC ? lenh cua cac nut Ribbon (F1 Xuat theo ngay, F4 Tao bang).
' Tuong ung src/commands/* cua ban Office.js.
'==============================================================================
Option Explicit

Private Const MAX_ISSUES_SHOWN As Long = 10
Private Const MSG_MAX_LENGTH As Long = 1000

' Hop thong bao chung cua add-in (MsgBox cat noi dung qua dai).
Public Sub ShowInfo(ByVal text As String, Optional ByVal kind As VbMsgBoxStyle = vbInformation)
    If Len(text) > MSG_MAX_LENGTH Then text = Left$(text, MSG_MAX_LENGTH - 1) & ChrW(8230)
    MsgBox text, kind, U("Nh\u1eadt k\u00fd TC")
End Sub

' Co workbook dang mo va sheet dang xem la worksheet (khong phai chart sheet).
Private Function HasWorksheet() As Boolean
    If ActiveWorkbook Is Nothing Then
        ShowInfo U("H\u00e3y m\u1edf m\u1ed9t file Excel tr\u01b0\u1edbc."), vbExclamation
    ElseIf TypeName(ActiveSheet) <> "Worksheet" Then
        ShowInfo U("H\u00e3y ch\u1ecdn m\u1ed9t sheet d\u1eef li\u1ec7u (kh\u00f4ng ph\u1ea3i sheet bi\u1ec3u \u0111\u1ed3)."), vbExclamation
    Else
        HasWorksheet = True
    End If
End Function

'------------------------------------------------------------------------------
' F4 Tao bang nhap tien do.
'------------------------------------------------------------------------------
Public Sub CmdTaoBang()
    Dim f As frmTaoBang, msg As String, createdAt As String, onNewSheet As Boolean, startCell As Range, rows As Long
    If Not HasWorksheet() Then Exit Sub
    Set f = New frmTaoBang
    f.Init ActiveSheet.Name, ActiveCell.Address(False, False)
    Do
        f.Show
        If Len(f.PickRequest) = 0 Then Exit Do
        f.ApplyPick PickRange(U("T\u1ea1o b\u1ea3ng"), U("Ch\u1ecdn \u00f4 b\u1eaft \u0111\u1ea7u c\u1ee7a b\u1ea3ng (g\u00f3c tr\u00ean-tr\u00e1i, \u00f4 ti\u00eau \u0111\u1ec1 STT):"), f.PickCurrent)
    Loop
    If f.Accepted Then
        onNewSheet = f.OnNewSheet
        rows = f.RowCount
        If Not onNewSheet Then Set startCell = f.StartCell
    End If
    Unload f
    Set f = Nothing
    If rows = 0 Then Exit Sub

    msg = CreateInputTable(onNewSheet, startCell, rows, createdAt)
    If Len(msg) > 0 Then ShowInfo msg, vbExclamation: Exit Sub
    ShowInfo U("\u0110\u00e3 t\u1ea1o b\u1ea3ng ") & createdAt & U(" g\u1ed3m ") & rows & U(" d\u00f2ng.") & vbLf & vbLf & _
        U("\u0110i\u1ec1n N\u1ed9i dung c\u00f4ng vi\u1ec7c, ng\u00e0y B\u1eaft \u0111\u1ea7u, ng\u00e0y K\u1ebft th\u00fac (dd/mm/yyyy). G\u00f5 ti\u1ebfp ngay d\u01b0\u1edbi d\u00f2ng cu\u1ed1i, b\u1ea3ng t\u1ef1 m\u1edf r\u1ed9ng.") & vbLf & _
        U("\u0110i\u1ec1n xong: \u0111\u1eb7t con tr\u1ecf trong b\u1ea3ng \u2192 b\u1ea5m ""Xu\u1ea5t theo ng\u00e0y"".")
End Sub

'------------------------------------------------------------------------------
' F1 Xuat nhat ky theo ngay.
'------------------------------------------------------------------------------
Public Sub CmdXuatNgay()
    Dim f As frmXuatNgay, source As Range, msg As String, initial As String, cc As Long, cs As Long, ce As Long
    If Not HasWorksheet() Then Exit Sub
    ' Goi y vung: vung dang chon; chon 1 o -> ca vung co du lieu cua sheet.
    Set source = ResolveDataRange(SelectedRange(), msg)
    If source Is Nothing Then
        If Not SelectedRange() Is Nothing Then initial = RefText(SelectedRange())
    Else
        initial = RefText(source)
    End If

    Set f = New frmXuatNgay
    f.Init initial
    Do
        f.Show
        If Len(f.PickRequest) = 0 Then Exit Do
        f.ApplyPick PickRange(U("Xu\u1ea5t nh\u1eadt k\u00fd theo ng\u00e0y"), _
            U("Ch\u1ecdn v\u00f9ng d\u1eef li\u1ec7u (g\u1ed3m d\u00f2ng ti\u00eau \u0111\u1ec1). Ch\u1ec9 ch\u1ecdn 1 \u00f4 = l\u1ea5y c\u1ea3 v\u00f9ng c\u00f3 d\u1eef li\u1ec7u c\u1ee7a sheet."), f.PickCurrent)
    Loop
    If f.Accepted Then
        Set source = f.Source
        cc = f.ColContent: cs = f.ColStart: ce = f.ColEnd
    Else
        Set source = Nothing
    End If
    Unload f
    Set f = Nothing
    If source Is Nothing Then Exit Sub
    ExportDailyLog source, cc, cs, ce
End Sub

' Doc lai gia tri moi nhat cua vung (CLAUDE.md ?4), trai theo ngay va tao sheet NK_TheoNgay.
Private Sub ExportDailyLog(ByVal source As Range, ByVal cc As Long, ByVal cs As Long, ByVal ce As Long)
    Dim values As Variant, date1904 As Boolean, wb As Workbook
    Dim items() As NKTC_WorkItem, itemCount As Long, issues() As NKTC_RowIssue, issueCount As Long, blankCount As Long
    Dim entries() As NKTC_DailyEntry, entryCount As Long, firstDay As Long, lastDay As Long
    Dim msg As String, issueText As String, sheetName As String, i As Long

    Set wb = source.Worksheet.Parent
    date1904 = wb.Date1904
    values = RangeGrid(source)
    ParseWorkItems values, cc, cs, ce, source.Row, date1904, items, itemCount, issues, issueCount, blankCount

    For i = 0 To issueCount - 1
        If i >= MAX_ISSUES_SHOWN Then
            issueText = issueText & vbLf & U("\u2026 v\u00e0 ") & (issueCount - MAX_ISSUES_SHOWN) & U(" d\u00f2ng kh\u00e1c.")
            Exit For
        End If
        issueText = issueText & vbLf & U("D\u00f2ng ") & issues(i).RowNumber & ": " & issues(i).Reason
    Next i

    msg = BuildDailyLog(items, itemCount, entries, entryCount, firstDay, lastDay)
    If Len(msg) > 0 Then ShowInfo msg & issueText, vbExclamation: Exit Sub

    sheetName = WriteDailyLogSheet(wb, entries, entryCount, date1904)
    msg = U("\u0110\u00e3 t\u1ea1o sheet """) & sheetName & U(""" g\u1ed3m ") & entryCount & U(" ng\u00e0y, t\u1eeb ") & FormatDay(firstDay) & _
        U(" \u0111\u1ebfn ") & FormatDay(lastDay) & "." & vbLf & itemCount & U(" c\u00f4ng vi\u1ec7c h\u1ee3p l\u1ec7.")
    If issueCount > 0 Then msg = msg & vbLf & vbLf & issueCount & U(" d\u00f2ng l\u1ed7i \u0111\u00e3 b\u1ecf qua:") & issueText
    msg = msg & vbLf & vbLf & U("C\u00f3 th\u1ec3 th\u00eam c\u1ed9t (Th\u1eddi ti\u1ebft, S\u1ed1 c\u00f4ng nh\u00e2n\u2026) v\u00e0o sheet n\u00e0y, r\u1ed3i d\u00f9ng ""Xu\u1ea5t v\u00e0o form"".")
    If issueCount > 0 Then ShowInfo msg, vbExclamation Else ShowInfo msg
End Sub
