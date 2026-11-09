Attribute VB_Name = "modNKTC_DailyLog"
'==============================================================================
' NHAT KY TC ? F1 Xuat nhat ky theo ngay: phan tinh toan (khong phu thuoc Excel).
' Tuong ung src/core/dailyLog/* cua ban Office.js.
'==============================================================================
Option Explicit

Public Const DAILY_SHEET_BASE_NAME As String = "NK_TheoNgay"
Public Const MAX_SPAN_DAYS As Long = 3660       ' R1.10: chan loi nhap nam
Public Const MAX_CELL_LENGTH As Long = 32767    ' so ky tu toi da cua mot o

' Bang nhap tien do (F4 Tao bang).
Public Const INPUT_SHEET_BASE_NAME As String = "TienDo"
Public Const DEFAULT_TABLE_ROWS As Long = 10
Public Const MAX_TABLE_ROWS As Long = 1000

' Mot cong viec hop le.
Public Type NKTC_WorkItem
    RowNumber As Long     ' so dong tren sheet (de bao loi)
    Content As String
    StartDay As Long
    EndDay As Long
End Type

' Dong bi bo qua va ly do (R1.6).
Public Type NKTC_RowIssue
    RowNumber As Long
    Reason As String
End Type

' Mot dong cua sheet nhat ky theo ngay.
Public Type NKTC_DailyEntry
    DayNum As Long
    Content As String     ' rong neu ngay khong co cong viec (R1.3)
    ItemCount As Long
End Type

'------------------------------------------------------------------------------
' Nhan dien cot (R1.12): vi tri 0-based trong dong tieu de; -1 neu khong nhan ra.
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

' headers: mang 1 chieu (0-based) gia tri dong tieu de. Ket qua qua tham so; -1 = chua nhan dien.
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
' Doc cong viec (R1.5, R1.6). values: mang 2 chieu Range.Value2 (1-based, gom dong tieu de o dong 1).
' firstRowNumber: so dong tren sheet cua dong tieu de. Cot 0-based trong vung.
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
                reason = U("Thi\u1ebfu n\u1ed9i dung c\u00f4ng vi\u1ec7c.")
            ElseIf s = 0 Then
                If IsBlankValue(rawStart) Then reason = U("Thi\u1ebfu ng\u00e0y b\u1eaft \u0111\u1ea7u.") Else reason = U("Ng\u00e0y b\u1eaft \u0111\u1ea7u kh\u00f4ng h\u1ee3p l\u1ec7.")
            Else
                ' R1.5: thieu ngay ket thuc -> cong viec 1 ngay.
                If IsBlankValue(rawEnd) Then e = s Else e = ToDayNumber(rawEnd, date1904)
                If e = 0 Then
                    reason = U("Ng\u00e0y k\u1ebft th\u00fac kh\u00f4ng h\u1ee3p l\u1ec7.")
                ElseIf e < s Then
                    reason = U("Ng\u00e0y k\u1ebft th\u00fac tr\u01b0\u1edbc ng\u00e0y b\u1eaft \u0111\u1ea7u.")
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
' Trai cong viec theo tung ngay va gop noi dung (R1.1?R1.4, R1.10).
' Tra ve chuoi loi; rong neu thanh cong.
'------------------------------------------------------------------------------
Public Function BuildDailyLog(ByRef items() As NKTC_WorkItem, ByVal itemCount As Long, _
                              ByRef entries() As NKTC_DailyEntry, ByRef entryCount As Long, _
                              ByRef firstDay As Long, ByRef lastDay As Long) As String
    Dim i As Long, d As Long, span As Long, idx As Long
    entryCount = 0
    If itemCount = 0 Then BuildDailyLog = U("Kh\u00f4ng c\u00f3 c\u00f4ng vi\u1ec7c h\u1ee3p l\u1ec7 \u0111\u1ec3 xu\u1ea5t."): Exit Function
    firstDay = items(0).StartDay: lastDay = items(0).EndDay
    For i = 1 To itemCount - 1
        If items(i).StartDay < firstDay Then firstDay = items(i).StartDay
        If items(i).EndDay > lastDay Then lastDay = items(i).EndDay
    Next i
    span = lastDay - firstDay + 1
    If span > MAX_SPAN_DAYS Then
        BuildDailyLog = U("Kho\u1ea3ng th\u1eddi gian t\u1eeb ") & FormatDay(firstDay) & U(" \u0111\u1ebfn ") & FormatDay(lastDay) & _
            " (" & span & U(" ng\u00e0y) v\u01b0\u1ee3t gi\u1edbi h\u1ea1n ") & MAX_SPAN_DAYS & U(" ng\u00e0y. Vui l\u00f2ng ki\u1ec3m tra l\u1ea1i ng\u00e0y nh\u1eadp.")
        Exit Function
    End If
    ReDim entries(0 To span - 1)
    For i = 0 To span - 1
        entries(i).DayNum = firstDay + i
    Next i
    ' Moi cong viec mot dong trong o, giu nguyen noi dung (R1.4).
    For i = 0 To itemCount - 1
        For d = items(i).StartDay To items(i).EndDay
            idx = d - firstDay
            If entries(idx).ItemCount > 0 Then entries(idx).Content = entries(idx).Content & vbLf
            entries(idx).Content = entries(idx).Content & items(i).Content
            entries(idx).ItemCount = entries(idx).ItemCount + 1
            If Len(entries(idx).Content) > MAX_CELL_LENGTH Then
                BuildDailyLog = U("N\u1ed9i dung ng\u00e0y ") & FormatDay(d) & U(" v\u01b0\u1ee3t qu\u00e1 ") & MAX_CELL_LENGTH & _
                    U(" k\u00fd t\u1ef1 (gi\u1edbi h\u1ea1n m\u1ed9t \u00f4 Excel).")
                Exit Function
            End If
        Next d
    Next i
    entryCount = span
End Function

' Ten sheet chua co: NK_TheoNgay, NK_TheoNgay_2, ... (R1.9; khong phan biet hoa/thuong).
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

' So dong cua bang nhap tien do: so nguyen 1..MAX_TABLE_ROWS; 0 neu khong hop le.
Public Function ParseRowCount(ByVal text As String) As Long
    Dim i As Long
    text = Trim$(text)
    If Len(text) = 0 Or Len(text) > 4 Then Exit Function
    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then Exit Function
    Next i
    If CLng(text) >= 1 And CLng(text) <= MAX_TABLE_ROWS Then ParseRowCount = CLng(text)
End Function

' Tieu de bang nhap tien do ? trung quy tac nhan dien cot cua F1 (R1.12).
Public Function InputTableHeaders() As Variant
    InputTableHeaders = Array("STT", U("N\u1ed9i dung c\u00f4ng vi\u1ec7c"), U("B\u1eaft \u0111\u1ea7u"), U("K\u1ebft th\u00fac"))
End Function
