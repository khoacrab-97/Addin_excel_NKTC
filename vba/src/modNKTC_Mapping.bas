Attribute VB_Name = "modNKTC_Mapping"
'==============================================================================
' NHẬT KÝ TC – F2 Xuất vào form: gán cột, danh sách ngày, chuyển ngày, kế hoạch ghi ô (không phụ thuộc Excel).
' Tương ứng src/core/mapping/* của bản Office.js.
'==============================================================================
Option Explicit

' Gán một cột nguồn (0-based trong vùng nguồn) vào một ô đích (R2.1).
Public Type NKTC_Mapping
    Column As Long
    SheetName As String
    Address As String      ' một ô, viết hoa, không có $ (ví dụ "E2")
End Type

' Một ngày điều hướng được và dòng tương ứng trong vùng nguồn.
Public Type NKTC_SourceDay
    DayNum As Long
    ValueRow As Long       ' chỉ số dòng trong mảng values (dòng tiêu đề = LBound)
    RowNumber As Long      ' số dòng trên sheet
End Type

' Kết quả một lệnh chuyển ngày.
Public Type NKTC_DayResolution
    Ok As Boolean
    TargetIndex As Long    ' vị trí trong danh sách ngày (khi Ok)
    Notice As String       ' thông báo kèm theo (khi Ok)
    IsInfo As Boolean      ' khi không Ok: True = đã ở đầu/cuối, False = lỗi
    Message As String      ' khi không Ok
End Type

Public Const CMD_FIRST As Long = 1
Public Const CMD_PREV As Long = 2
Public Const CMD_NEXT As Long = 3
Public Const CMD_LAST As Long = 4

'------------------------------------------------------------------------------
' Khóa ô đích
'------------------------------------------------------------------------------
Public Function TargetKey(ByVal sheetName As String, ByVal address As String) As String
    TargetKey = LCase$(sheetName) & "!" & UCase$(address)
End Function

Public Function FormatTarget(ByVal sheetName As String, ByVal address As String) As String
    FormatTarget = sheetName & "!" & address
End Function

' Khóa của bộ ô đích (không phụ thuộc thứ tự gán) — dùng cho xác nhận ghi đè (R2.5).
Public Function MappingKey(ByRef maps() As NKTC_Mapping, ByVal count As Long) As String
    Dim keys() As String, i As Long, j As Long, tmp As String
    If count = 0 Then Exit Function
    ReDim keys(0 To count - 1)
    For i = 0 To count - 1
        keys(i) = TargetKey(maps(i).SheetName, maps(i).Address)
    Next i
    ' Sắp xếp theo mã ký tự (giống bản Office.js).
    For i = 1 To count - 1
        tmp = keys(i)
        j = i - 1
        Do While j >= 0
            If StrComp(keys(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i
    MappingKey = Join(keys, "|")
End Function

'------------------------------------------------------------------------------
' Cột Ngày và danh sách ngày (R2.10, R2.13). values: Range.Value2 (2 chiều), dòng đầu là tiêu đề.
'------------------------------------------------------------------------------
' Cột Ngày (0-based): tiêu đề "Ngày" -> tiêu đề có từ "ngay" -> cột đầu tiên toàn ngày hợp lệ; -1 nếu không có.
Public Function DetectDateColumn(ByVal values As Variant) As Long
    Dim r0 As Long, c0 As Long, nCols As Long, c As Long, r As Long, h As String, filled As Long, ok As Boolean
    DetectDateColumn = -1
    r0 = LBound(values, 1): c0 = LBound(values, 2)
    nCols = UBound(values, 2) - c0 + 1
    For c = 0 To nCols - 1
        If VarType(values(r0, c0 + c)) = vbString Then
            If NormalizeHeader(values(r0, c0 + c)) = "ngay" Then DetectDateColumn = c: Exit Function
        End If
    Next c
    For c = 0 To nCols - 1
        If VarType(values(r0, c0 + c)) = vbString Then
            h = " " & NormalizeHeader(values(r0, c0 + c)) & " "
            If InStr(1, h, " ngay ", vbBinaryCompare) > 0 Then DetectDateColumn = c: Exit Function
        End If
    Next c
    For c = 0 To nCols - 1
        filled = 0: ok = True
        For r = r0 + 1 To UBound(values, 1)
            If Not IsBlankValue(values(r, c0 + c)) Then
                filled = filled + 1
                If ToDayNumber(values(r, c0 + c)) = 0 Then ok = False: Exit For
            End If
        Next r
        If ok And filled > 0 Then DetectDateColumn = c: Exit Function
    Next c
End Function

' Danh sách ngày tăng dần, không trùng. Dòng có ô Ngày trống bỏ qua; ngày sai hoặc trùng ghi vào issues.
Public Sub BuildDayIndex(ByVal values As Variant, ByVal dateColumn As Long, ByVal firstRowNumber As Long, _
                         ByVal date1904 As Boolean, ByRef days() As NKTC_SourceDay, ByRef dayCount As Long, _
                         ByRef issues As Collection)
    Dim r0 As Long, c0 As Long, r As Long, d As Long, i As Long, j As Long, raw As Variant
    Dim seen As Collection, rowNumber As Long, firstRow As Variant, tmp As NKTC_SourceDay
    Set issues = New Collection
    Set seen = New Collection
    dayCount = 0
    ReDim days(0 To 0)
    r0 = LBound(values, 1): c0 = LBound(values, 2)
    For r = r0 + 1 To UBound(values, 1)
        rowNumber = firstRowNumber + (r - r0)
        raw = values(r, c0 + dateColumn)
        If Not IsBlankValue(raw) Then
            d = ToDayNumber(raw, date1904)
            If d = 0 Then
                issues.Add "Dòng " & rowNumber & ": Ngày không hợp lệ."
            Else
                firstRow = Empty
                On Error Resume Next
                firstRow = seen("d" & d)
                On Error GoTo 0
                If Not IsEmpty(firstRow) Then
                    issues.Add "Dòng " & rowNumber & ": Trùng ngày " & FormatDay(d) & " với dòng " & firstRow & "; bỏ qua dòng này."
                Else
                    seen.Add rowNumber, "d" & d
                    ReDim Preserve days(0 To dayCount)
                    days(dayCount).DayNum = d
                    days(dayCount).ValueRow = r
                    days(dayCount).RowNumber = rowNumber
                    dayCount = dayCount + 1
                End If
            End If
        End If
    Next r
    ' Sắp xếp tăng dần theo ngày (chèn trực tiếp; vùng nguồn thường đã theo thứ tự).
    For i = 1 To dayCount - 1
        tmp = days(i)
        j = i - 1
        Do While j >= 0
            If days(j).DayNum <= tmp.DayNum Then Exit Do
            days(j + 1) = days(j)
            j = j - 1
        Loop
        days(j + 1) = tmp
    Next i
End Sub

Public Function IndexOfDay(ByRef days() As NKTC_SourceDay, ByVal dayCount As Long, ByVal dayNum As Long) As Long
    Dim i As Long
    IndexOfDay = -1
    For i = 0 To dayCount - 1
        If days(i).DayNum = dayNum Then IndexOfDay = i: Exit Function
    Next i
End Function

' Ngày cần hiển thị cho một lệnh chuyển ngày. currentDay = 0 nếu chưa hiển thị ngày nào.
Public Function ResolveDay(ByRef days() As NKTC_SourceDay, ByVal dayCount As Long, ByVal currentDay As Long, _
                           ByVal command As Long) As NKTC_DayResolution
    Dim res As NKTC_DayResolution, idx As Long, nextIdx As Long
    If dayCount = 0 Then
        res.Message = "Vùng nguồn không có ngày hợp lệ ở cột Ngày đã chọn."
        ResolveDay = res: Exit Function
    End If
    Select Case command
        Case CMD_FIRST: res.Ok = True: res.TargetIndex = 0
        Case CMD_LAST: res.Ok = True: res.TargetIndex = dayCount - 1
        Case CMD_PREV, CMD_NEXT
            idx = -1
            If currentDay > 0 Then idx = IndexOfDay(days, dayCount, currentDay)
            If idx = -1 Then
                res.Ok = True: res.TargetIndex = 0
                If currentDay > 0 Then
                    res.Notice = "Không còn ngày " & FormatDay(currentDay) & " trong vùng nguồn; đã hiển thị ngày " & _
                        FormatDay(days(0).DayNum) & "."
                End If
            Else
                If command = CMD_NEXT Then nextIdx = idx + 1 Else nextIdx = idx - 1
                If nextIdx < 0 Then
                    res.IsInfo = True
                    res.Message = "Đang ở ngày đầu tiên (" & FormatDay(days(0).DayNum) & ")."
                ElseIf nextIdx > dayCount - 1 Then
                    res.IsInfo = True
                    res.Message = "Đang ở ngày cuối cùng (" & FormatDay(days(dayCount - 1).DayNum) & ")."
                Else
                    res.Ok = True: res.TargetIndex = nextIdx
                End If
            End If
    End Select
    ResolveDay = res
End Function

'------------------------------------------------------------------------------
' Khớp cột theo tên tiêu đề (R2.17): không phân biệt hoa/thường, khoảng trắng thừa; trùng tên theo thứ tự.
'------------------------------------------------------------------------------
Private Function HeaderKeys(ByRef headers() As String) As String()
    Dim keys() As String, counts As Collection, i As Long, name As String, cnt As Variant
    Set counts = New Collection
    ReDim keys(LBound(headers) To UBound(headers))
    For i = LBound(headers) To UBound(headers)
        name = NormalizeHeaderName(headers(i))
        cnt = Empty
        On Error Resume Next
        cnt = counts("k" & name)
        On Error GoTo 0
        If IsEmpty(cnt) Then
            cnt = 0
        Else
            counts.Remove "k" & name
        End If
        counts.Add cnt + 1, "k" & name
        If Len(name) > 0 Then keys(i) = cnt & "|" & name   ' tiêu đề trống không khớp
    Next i
    HeaderKeys = keys
End Function

' Vị trí mới (0-based trong newHeaders) của từng cột cũ; -1 nếu không còn (bị xóa/đổi tên).
' oldHeaders/newHeaders: mảng 0-based.
Public Function MatchHeaders(ByRef oldHeaders() As String, ByRef newHeaders() As String) As Long()
    Dim oldKeys() As String, newKeys() As String, result() As Long, i As Long, j As Long
    oldKeys = HeaderKeys(oldHeaders)
    newKeys = HeaderKeys(newHeaders)
    ReDim result(LBound(oldHeaders) To UBound(oldHeaders))
    For i = LBound(oldKeys) To UBound(oldKeys)
        result(i) = -1
        If Len(oldKeys(i)) > 0 Then
            For j = LBound(newKeys) To UBound(newKeys)
                If newKeys(j) = oldKeys(i) Then result(i) = j: Exit For
            Next j
        End If
    Next i
    MatchHeaders = result
End Function

'------------------------------------------------------------------------------
' Ghi ô (R2.6, R2.11, R2.15)
'------------------------------------------------------------------------------
' Giá trị ghi vào ô: chuỗi khác rỗng thêm dấu ' để Excel giữ dạng văn bản; số/logic giữ nguyên; còn lại rỗng.
Public Function ToCellInput(ByVal v As Variant) As Variant
    Select Case VarType(v)
        Case vbString
            If Len(v) = 0 Then ToCellInput = "" Else ToCellInput = "'" & v
        Case vbBoolean, vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            ToCellInput = v
        Case Else
            ToCellInput = ""
    End Select
End Function

' Giá trị ghi cho một gán: cột Ngày vào ô đang chứa mẫu "Ngày … tháng … năm …" -> chữ đã rải ngày.
' currentText: chữ hiện có của ô đích (chuỗi rỗng nếu ô không chứa chữ).
Public Function WriteValue(ByVal sourceValue As Variant, ByVal isDateColumn As Boolean, ByVal currentText As String, _
                           Optional ByVal date1904 As Boolean = False) As Variant
    Dim d As Long, filled As String
    If isDateColumn And Len(currentText) > 0 Then
        d = ToDayNumber(sourceValue, date1904)
        If d > 0 Then filled = FillDatePattern(currentText, d)
        If Len(filled) > 0 Then WriteValue = ToCellInput(filled): Exit Function
    End If
    WriteValue = ToCellInput(sourceValue)
End Function

' Kiểm tra bộ gán: có ít nhất một gán, cột trong vùng nguồn, mỗi ô chỉ nhận một cột. Trả về lỗi (rỗng nếu hợp lệ).
Public Function ValidateMappings(ByRef maps() As NKTC_Mapping, ByVal count As Long, ByVal columnCount As Long) As Collection
    Dim errors As New Collection, owners As New Collection, i As Long, key As String, owner As Variant
    If count = 0 Then errors.Add "Chưa gán cột nào vào ô của form mẫu."
    For i = 0 To count - 1
        If maps(i).Column < 0 Or maps(i).Column >= columnCount Then
            errors.Add "Cột thứ " & (maps(i).Column + 1) & " không nằm trong vùng nguồn."
        Else
            key = TargetKey(maps(i).SheetName, maps(i).Address)
            owner = Empty
            On Error Resume Next
            owner = owners(key)
            On Error GoTo 0
            If IsEmpty(owner) Then
                owners.Add maps(i).Column, key
            ElseIf owner <> maps(i).Column Then
                errors.Add "Ô " & FormatTarget(maps(i).SheetName, maps(i).Address) & " đang được gán cho 2 cột khác nhau."
            End If
        End If
    Next i
    Set ValidateMappings = errors
End Function
