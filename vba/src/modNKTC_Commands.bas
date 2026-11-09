Attribute VB_Name = "modNKTC_Commands"
'==============================================================================
' NHẬT KÝ TC – lệnh của các nút Ribbon (F1 Xuất theo ngày, F4 Tạo bảng).
' Tương ứng src/commands/* của bản Office.js.
'==============================================================================
Option Explicit

Private Const MAX_ISSUES_SHOWN As Long = 10

' Có workbook đang mở và sheet đang xem là worksheet (không phải chart sheet).
Private Function HasWorksheet() As Boolean
    If ActiveWorkbook Is Nothing Then
        ShowInfo "Hãy mở một file Excel trước.", vbExclamation
    ElseIf TypeName(ActiveSheet) <> "Worksheet" Then
        ShowInfo "Hãy chọn một sheet dữ liệu (không phải sheet biểu đồ).", vbExclamation
    Else
        HasWorksheet = True
    End If
End Function

'------------------------------------------------------------------------------
' F4 Tạo bảng nhập tiến độ.
'------------------------------------------------------------------------------
Public Sub CmdTaoBang()
    Dim f As frmTaoBang, msg As String, createdAt As String, onNewSheet As Boolean, startCell As Range, rows As Long
    If Not HasWorksheet() Then Exit Sub
    Set f = New frmTaoBang
    f.Init ActiveSheet.Name, ActiveCell.Address(False, False)
    Do
        f.Show
        If Len(f.PickRequest) = 0 Then Exit Do
        f.ApplyPick PickRange("Tạo bảng", "Chọn ô bắt đầu của bảng (góc trên-trái, ô tiêu đề STT):", f.PickCurrent)
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
    ShowInfo "Đã tạo bảng " & createdAt & " gồm " & rows & " dòng." & vbLf & vbLf & _
        "Điền Nội dung công việc, ngày Bắt đầu, ngày Kết thúc (dd/mm/yyyy). Gõ tiếp ngay dưới dòng cuối, bảng tự mở rộng." & vbLf & _
        "Điền xong: đặt con trỏ trong bảng → bấm ""Xuất theo ngày""."
End Sub

'------------------------------------------------------------------------------
' F1 Xuất nhật ký theo ngày.
'------------------------------------------------------------------------------
Public Sub CmdXuatNgay()
    Dim f As frmXuatNgay, source As Range, msg As String, initial As String, cc As Long, cs As Long, ce As Long
    If Not HasWorksheet() Then Exit Sub
    ' Gợi ý vùng: vùng đang chọn; chọn 1 ô -> cả vùng có dữ liệu của sheet.
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
        f.ApplyPick PickRange("Xuất nhật ký theo ngày", _
            "Chọn vùng dữ liệu (gồm dòng tiêu đề). Chỉ chọn 1 ô = lấy cả vùng có dữ liệu của sheet.", f.PickCurrent)
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

' Đọc lại giá trị mới nhất của vùng (CLAUDE.md §4), trải theo ngày và tạo sheet NK_TheoNgay.
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
            issueText = issueText & vbLf & "… và " & (issueCount - MAX_ISSUES_SHOWN) & " dòng khác."
            Exit For
        End If
        issueText = issueText & vbLf & "Dòng " & issues(i).RowNumber & ": " & issues(i).Reason
    Next i

    msg = BuildDailyLog(items, itemCount, entries, entryCount, firstDay, lastDay)
    If Len(msg) > 0 Then ShowInfo msg & issueText, vbExclamation: Exit Sub

    sheetName = WriteDailyLogSheet(wb, entries, entryCount, date1904)
    msg = "Đã tạo sheet """ & sheetName & """ gồm " & entryCount & " ngày, từ " & FormatDay(firstDay) & _
        " đến " & FormatDay(lastDay) & "." & vbLf & itemCount & " công việc hợp lệ."
    If issueCount > 0 Then msg = msg & vbLf & vbLf & issueCount & " dòng lỗi đã bỏ qua:" & issueText
    msg = msg & vbLf & vbLf & "Có thể thêm cột (Thời tiết, Số công nhân…) vào sheet này, rồi dùng ""Xuất vào form""."
    If issueCount > 0 Then ShowInfo msg, vbExclamation Else ShowInfo msg
End Sub
