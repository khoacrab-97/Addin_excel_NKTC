Attribute VB_Name = "modNKTC_Mapping"
'==============================================================================
' NHAT KY TC ? F2 Xuat vao form: gan cot, danh sach ngay, chuyen ngay, ke hoach ghi o (khong phu thuoc Excel).
' Tuong ung src/core/mapping/* cua ban Office.js.
'==============================================================================
Option Explicit

' Gan mot cot nguon (0-based trong vung nguon) vao mot o dich (R2.1).
Public Type NKTC_Mapping
    Column As Long
    SheetName As String
    Address As String      ' mot o, viet hoa, khong co $ (vi du "E2")
End Type

' Mot ngay dieu huong duoc va dong tuong ung trong vung nguon.
Public Type NKTC_SourceDay
    DayNum As Long
    ValueRow As Long       ' chi so dong trong mang values (dong tieu de = LBound)
    RowNumber As Long      ' so dong tren sheet
End Type

' Ket qua mot lenh chuyen ngay.
Public Type NKTC_DayResolution
    Ok As Boolean
    TargetIndex As Long    ' vi tri trong danh sach ngay (khi Ok)
    Notice As String       ' thong bao kem theo (khi Ok)
    IsInfo As Boolean      ' khi khong Ok: True = da o dau/cuoi, False = loi
    Message As String      ' khi khong Ok
End Type

Public Const CMD_FIRST As Long = 1
Public Const CMD_PREV As Long = 2
Public Const CMD_NEXT As Long = 3
Public Const CMD_LAST As Long = 4
Public Const CMD_GOTO As Long = 5      ' mot ngay cu the (targetDay)

'------------------------------------------------------------------------------
' Khoa o dich
'------------------------------------------------------------------------------
Public Function TargetKey(ByVal sheetName As String, ByVal address As String) As String
    TargetKey = LCase$(sheetName) & "!" & UCase$(address)
End Function

Public Function FormatTarget(ByVal sheetName As String, ByVal address As String) As String
    FormatTarget = sheetName & "!" & address
End Function

' Khoa cua bo o dich (khong phu thuoc thu tu gan) ? dung cho xac nhan ghi de (R2.5).
Public Function MappingKey(ByRef maps() As NKTC_Mapping, ByVal count As Long) As String
    Dim keys() As String, i As Long, j As Long, tmp As String
    If count = 0 Then Exit Function
    ReDim keys(0 To count - 1)
    For i = 0 To count - 1
        keys(i) = TargetKey(maps(i).SheetName, maps(i).Address)
    Next i
    ' Sap xep theo ma ky tu (giong ban Office.js).
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
' Cot Ngay va danh sach ngay (R2.10, R2.13). values: Range.Value2 (2 chieu), dong dau la tieu de.
'------------------------------------------------------------------------------
' Cot Ngay (0-based): tieu de "Ngay" -> tieu de co tu "ngay" -> cot dau tien toan ngay hop le; -1 neu khong co.
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

' Danh sach ngay tang dan, khong trung. Dong co o Ngay trong bo qua; ngay sai hoac trung ghi vao issues.
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
                issues.Add U("D\u00f2ng ") & rowNumber & U(": Ng\u00e0y kh\u00f4ng h\u1ee3p l\u1ec7.")
            Else
                firstRow = Empty
                On Error Resume Next
                firstRow = seen("d" & d)
                On Error GoTo 0
                If Not IsEmpty(firstRow) Then
                    issues.Add U("D\u00f2ng ") & rowNumber & U(": Tr\u00f9ng ng\u00e0y ") & FormatDay(d) & U(" v\u1edbi d\u00f2ng ") & firstRow & U("; b\u1ecf qua d\u00f2ng n\u00e0y.")
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
    ' Sap xep tang dan theo ngay (chen truc tiep; vung nguon thuong da theo thu tu).
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

' Ngay can hien thi cho mot lenh chuyen ngay. currentDay = 0 neu chua hien thi ngay nao.
Public Function ResolveDay(ByRef days() As NKTC_SourceDay, ByVal dayCount As Long, ByVal currentDay As Long, _
                           ByVal command As Long, Optional ByVal targetDay As Long = 0) As NKTC_DayResolution
    Dim res As NKTC_DayResolution, idx As Long, nextIdx As Long
    If dayCount = 0 Then
        res.Message = U("V\u00f9ng ngu\u1ed3n kh\u00f4ng c\u00f3 ng\u00e0y h\u1ee3p l\u1ec7 \u1edf c\u1ed9t Ng\u00e0y \u0111\u00e3 ch\u1ecdn.")
        ResolveDay = res: Exit Function
    End If
    Select Case command
        Case CMD_FIRST: res.Ok = True: res.TargetIndex = 0
        Case CMD_LAST: res.Ok = True: res.TargetIndex = dayCount - 1
        Case CMD_GOTO
            idx = IndexOfDay(days, dayCount, targetDay)
            If idx >= 0 Then
                res.Ok = True: res.TargetIndex = idx
            Else
                res.Message = U("Kh\u00f4ng c\u00f3 ng\u00e0y ") & FormatDay(targetDay) & U(" trong v\u00f9ng ngu\u1ed3n.")
            End If
        Case CMD_PREV, CMD_NEXT
            idx = -1
            If currentDay > 0 Then idx = IndexOfDay(days, dayCount, currentDay)
            If idx = -1 Then
                res.Ok = True: res.TargetIndex = 0
                If currentDay > 0 Then
                    res.Notice = U("Kh\u00f4ng c\u00f2n ng\u00e0y ") & FormatDay(currentDay) & U(" trong v\u00f9ng ngu\u1ed3n; \u0111\u00e3 hi\u1ec3n th\u1ecb ng\u00e0y ") & _
                        FormatDay(days(0).DayNum) & "."
                End If
            Else
                If command = CMD_NEXT Then nextIdx = idx + 1 Else nextIdx = idx - 1
                If nextIdx < 0 Then
                    res.IsInfo = True
                    res.Message = U("\u0110ang \u1edf ng\u00e0y \u0111\u1ea7u ti\u00ean (") & FormatDay(days(0).DayNum) & ")."
                ElseIf nextIdx > dayCount - 1 Then
                    res.IsInfo = True
                    res.Message = U("\u0110ang \u1edf ng\u00e0y cu\u1ed1i c\u00f9ng (") & FormatDay(days(dayCount - 1).DayNum) & ")."
                Else
                    res.Ok = True: res.TargetIndex = nextIdx
                End If
            End If
    End Select
    ResolveDay = res
End Function

'------------------------------------------------------------------------------
' Khop cot theo ten tieu de (R2.17): khong phan biet hoa/thuong, khoang trang thua; trung ten theo thu tu.
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
        If Len(name) > 0 Then keys(i) = cnt & "|" & name   ' tieu de trong khong khop
    Next i
    HeaderKeys = keys
End Function

' Vi tri moi (0-based trong newHeaders) cua tung cot cu; -1 neu khong con (bi xoa/doi ten).
' oldHeaders/newHeaders: mang 0-based.
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
' Ghi o (R2.6, R2.11, R2.15)
'------------------------------------------------------------------------------
' Gia tri ghi vao o: chuoi khac rong them dau ' de Excel giu dang van ban; so/logic giu nguyen; con lai rong.
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

' Gia tri ghi cho mot gan: cot Ngay vao o dang chua mau "Ngay ? thang ? nam ?" -> chu da rai ngay.
' currentText: chu hien co cua o dich (chuoi rong neu o khong chua chu).
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

' Kiem tra bo gan: co it nhat mot gan, cot trong vung nguon, moi o chi nhan mot cot. Tra ve loi (rong neu hop le).
Public Function ValidateMappings(ByRef maps() As NKTC_Mapping, ByVal count As Long, ByVal columnCount As Long) As Collection
    Dim errors As New Collection, owners As New Collection, i As Long, key As String, owner As Variant
    If count = 0 Then errors.Add U("Ch\u01b0a g\u00e1n c\u1ed9t n\u00e0o v\u00e0o \u00f4 c\u1ee7a form m\u1eabu.")
    For i = 0 To count - 1
        If maps(i).Column < 0 Or maps(i).Column >= columnCount Then
            errors.Add U("C\u1ed9t th\u1ee9 ") & (maps(i).Column + 1) & U(" kh\u00f4ng n\u1eb1m trong v\u00f9ng ngu\u1ed3n.")
        Else
            key = TargetKey(maps(i).SheetName, maps(i).Address)
            owner = Empty
            On Error Resume Next
            owner = owners(key)
            On Error GoTo 0
            If IsEmpty(owner) Then
                owners.Add maps(i).Column, key
            ElseIf owner <> maps(i).Column Then
                errors.Add U("\u00d4 ") & FormatTarget(maps(i).SheetName, maps(i).Address) & U(" \u0111ang \u0111\u01b0\u1ee3c g\u00e1n cho 2 c\u1ed9t kh\u00e1c nhau.")
            End If
        End If
    Next i
    Set ValidateMappings = errors
End Function
