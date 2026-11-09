Attribute VB_Name = "modNKTC_DailyLog"
'==============================================================================
' NHẬT KÝ TC – F1 Xuất nhật ký theo ngày: phần tính toán (không phụ thuộc Excel).
' Tương ứng src/core/dailyLog/* của bản Office.js.
'==============================================================================
Option Explicit

Public Const DAILY_SHEET_BASE_NAME As String = "NK_TheoNgay"
Public Const MAX_SPAN_DAYS As Long = 3660       ' R1.10: chặn lỗi nhập năm
Public Const MAX_CELL_LENGTH As Long = 32767    ' số ký tự tối đa của một ô

' Một công việc hợp lệ.
Public Type NKTC_WorkItem
    RowNumber As Long     ' số dòng trên sheet (để báo lỗi)
    Content As String
    StartDay As Long
    EndDay As Long
End Type

' Dòng bị bỏ qua và lý do (R1.6).
Public Type NKTC_RowIssue
    RowNumber As Long
    Reason As String
End Type

' Một dòng của sheet nhật ký theo ngày.
Public Type NKTC_DailyEntry
    DayNum As Long
    Content As String     ' rỗng nếu ngày không có công việc (R1.3)
    ItemCount As Long
End Type

'------------------------------------------------------------------------------
' Nhận diện cột (R1.12): vị trí 0-based trong dòng tiêu đề; -1 nếu không nhận ra.
'------------------------------------------------------------------------------
Private Function ExactList(ByVal key As String) As Variant
    Select Case key
        Case "content"
            ExactList = Array("noi dung cong viec", "ten cong viec", "cong viec", "ten hang muc cong tac", _
                "ten hang muc", "hang muc", "ten cong tac", "cong tac", "noi dung")
        Case "start": ExactList = Array("ngay bat dau", "bat dau", "tu ngay")
        Case "end": ExactList = Array("ngay ket thuc", "ket thuc", "den ngay")
    End Select
End Function

Private Function ContainsList(ByVal key As String) As Variant
    Select Case key
        Case "content": ContainsList = Array("cong viec", "cong tac", "hang muc")
        Case "start": ContainsList = Array("bat dau")
        Case "end": ContainsList = Array("ket thuc")
    End Select
End Function

Private Function InList(ByVal s As String, ByVal list As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(list)
        If s = list(i) Then InList = True: Exit Function
    Next i
End Function

Private Function ContainsAny(ByVal s As String, ByVal list As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(list)
        If InStr(1, s, list(i), vbBinaryCompare) > 0 Then ContainsAny = True: Exit Function
    Next i
End Function

' headers: mảng 1 chiều (0-based) giá trị dòng tiêu đề. Kết quả qua tham số; -1 = chưa nhận diện.
Public Sub DetectColumns(ByVal headers As Variant, ByRef colContent As Long, ByRef colStart As Long, ByRef colEnd As Long)
    Dim n As Long, i As Long, norm() As String, used() As Boolean, keys As Variant, k As Long, found As Long, pass As Long
    colContent = -1: colStart = -1: colEnd = -1
    n = UBound(headers) - LBound(headers) + 1
    If n <= 0 Then Exit Sub
    ReDim norm(0 To n - 1)
    ReDim used(0 To n - 1)
    For i = 0 To n - 1
        If VarType(headers(LBound(headers) + i)) = vbString Then norm(i) = NormalizeHeader(headers(LBound(headers) + i))
    Next i
    keys = Array("start", "end", "content")
    For k = 0 To 2
        found = -1
        For pass = 1 To 2
            For i = 0 To n - 1
                If Not used(i) And Len(norm(i)) > 0 Then
                    If pass = 1 Then
                        If InList(norm(i), ExactList(keys(k))) Then found = i
                    Else
                        If ContainsAny(norm(i), ContainsList(keys(k))) Then found = i
                    End If
                    If found >= 0 Then Exit For
                End If
            Next i
            If found >= 0 Then Exit For
        Next pass
        If found >= 0 Then
            used(found) = True
            Select Case keys(k)
                Case "content": colContent = found
                Case "start": colStart = found
                Case "end": colEnd = found
            End Select
        End If
    Next k
End Sub

Public Function IsCompleteColumns(ByVal colContent As Long, ByVal colStart As Long, ByVal colEnd As Long) As Boolean
    If colContent < 0 Or colStart < 0 Or colEnd < 0 Then Exit Function
    IsCompleteColumns = (colContent <> colStart And colContent <> colEnd And colStart <> colEnd)
End Function

'------------------------------------------------------------------------------
' Đọc công việc (R1.5, R1.6). values: mảng 2 chiều Range.Value2 (1-based, gồm dòng tiêu đề ở dòng 1).
' firstRowNumber: số dòng trên sheet của dòng tiêu đề. Cột 0-based trong vùng.
'------------------------------------------------------------------------------
Private Function ContentText(ByVal v As Variant) As String
    If VarType(v) = vbString Then ContentText = Trim$(v) Else ContentText = CellText(v)
End Function

Public Sub ParseWorkItems(ByVal values As Variant, ByVal colContent As Long, ByVal colStart As Long, ByVal colEnd As Long, _
                          ByVal firstRowNumber As Long, ByVal date1904 As Boolean, _
                          ByRef items() As NKTC_WorkItem, ByRef itemCount As Long, _
                          ByRef issues() As NKTC_RowIssue, ByRef issueCount As Long, ByRef blankCount As Long)
    Dim r As Long, r0 As Long, c0 As Long, rawContent As Variant, rawStart As Variant, rawEnd As Variant
    Dim content As String, s As Long, e As Long, rowNumber As Long, reason As String
    itemCount = 0: issueCount = 0: blankCount = 0
    r0 = LBound(values, 1): c0 = LBound(values, 2)
    ReDim items(0 To 0): ReDim issues(0 To 0)
    For r = r0 + 1 To UBound(values, 1)
        rowNumber = firstRowNumber + (r - r0)
        rawContent = values(r, c0 + colContent)
        rawStart = values(r, c0 + colStart)
        rawEnd = values(r, c0 + colEnd)
        reason = ""
        If IsBlankValue(rawContent) And IsBlankValue(rawStart) And IsBlankValue(rawEnd) Then
            blankCount = blankCount + 1
        Else
            content = ContentText(rawContent)
            s = ToDayNumber(rawStart, date1904)
            If Len(content) = 0 Then
                reason = "Thiếu nội dung công việc."
            ElseIf s = 0 Then
                If IsBlankValue(rawStart) Then reason = "Thiếu ngày bắt đầu." Else reason = "Ngày bắt đầu không hợp lệ."
            Else
                ' R1.5: thiếu ngày kết thúc -> công việc 1 ngày.
                If IsBlankValue(rawEnd) Then e = s Else e = ToDayNumber(rawEnd, date1904)
                If e = 0 Then
                    reason = "Ngày kết thúc không hợp lệ."
                ElseIf e < s Then
                    reason = "Ngày kết thúc trước ngày bắt đầu."
                End If
            End If
            If Len(reason) > 0 Then
                ReDim Preserve issues(0 To issueCount)
                issues(issueCount).RowNumber = rowNumber
                issues(issueCount).Reason = reason
                issueCount = issueCount + 1
            Else
                ReDim Preserve items(0 To itemCount)
                items(itemCount).RowNumber = rowNumber
                items(itemCount).Content = content
                items(itemCount).StartDay = s
                items(itemCount).EndDay = e
                itemCount = itemCount + 1
            End If
        End If
    Next r
End Sub

'------------------------------------------------------------------------------
' Trải công việc theo từng ngày và gộp nội dung (R1.1–R1.4, R1.10).
' Trả về chuỗi lỗi; rỗng nếu thành công.
'------------------------------------------------------------------------------
Public Function BuildDailyLog(ByRef items() As NKTC_WorkItem, ByVal itemCount As Long, _
                              ByRef entries() As NKTC_DailyEntry, ByRef entryCount As Long, _
                              ByRef firstDay As Long, ByRef lastDay As Long) As String
    Dim i As Long, d As Long, span As Long, idx As Long
    entryCount = 0
    If itemCount = 0 Then BuildDailyLog = "Không có công việc hợp lệ để xuất.": Exit Function
    firstDay = items(0).StartDay: lastDay = items(0).EndDay
    For i = 1 To itemCount - 1
        If items(i).StartDay < firstDay Then firstDay = items(i).StartDay
        If items(i).EndDay > lastDay Then lastDay = items(i).EndDay
    Next i
    span = lastDay - firstDay + 1
    If span > MAX_SPAN_DAYS Then
        BuildDailyLog = "Khoảng thời gian từ " & FormatDay(firstDay) & " đến " & FormatDay(lastDay) & _
            " (" & span & " ngày) vượt giới hạn " & MAX_SPAN_DAYS & " ngày. Vui lòng kiểm tra lại ngày nhập."
        Exit Function
    End If
    ReDim entries(0 To span - 1)
    For i = 0 To span - 1
        entries(i).DayNum = firstDay + i
    Next i
    ' Mỗi công việc một dòng trong ô, giữ nguyên nội dung (R1.4).
    For i = 0 To itemCount - 1
        For d = items(i).StartDay To items(i).EndDay
            idx = d - firstDay
            If entries(idx).ItemCount > 0 Then entries(idx).Content = entries(idx).Content & vbLf
            entries(idx).Content = entries(idx).Content & items(i).Content
            entries(idx).ItemCount = entries(idx).ItemCount + 1
            If Len(entries(idx).Content) > MAX_CELL_LENGTH Then
                BuildDailyLog = "Nội dung ngày " & FormatDay(d) & " vượt quá " & MAX_CELL_LENGTH & _
                    " ký tự (giới hạn một ô Excel)."
                Exit Function
            End If
        Next d
    Next i
    entryCount = span
End Function

' Tên sheet chưa có: NK_TheoNgay, NK_TheoNgay_2, ... (R1.9; không phân biệt hoa/thường).
Public Function UniqueSheetName(ByVal existing As Collection, Optional ByVal base As String = DAILY_SHEET_BASE_NAME) As String
    Dim suffix As Long, candidate As String
    candidate = base
    suffix = 1
    Do While NameTaken(existing, candidate)
        suffix = suffix + 1
        candidate = base & "_" & suffix
    Loop
    UniqueSheetName = candidate
End Function

Private Function NameTaken(ByVal existing As Collection, ByVal name As String) As Boolean
    Dim v As Variant
    For Each v In existing
        If LCase$(v) = LCase$(name) Then NameTaken = True: Exit Function
    Next v
End Function
