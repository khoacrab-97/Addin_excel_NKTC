Attribute VB_Name = "modNKTC_Date"
'==============================================================================
' NHẬT KÝ TC – ngày tháng (không phụ thuộc Excel).
' Ngày dùng số nguyên theo hệ ngày 1900 của Excel (DayNumber, kiểu Long), ví dụ 46327 = 01/11/2026.
' Ngày của VBA (Date) trùng hệ 1900 của Excel từ 01/03/1900 trở đi.
'==============================================================================
Option Explicit

Public Const MIN_DAY As Long = 61            ' 01/03/1900
Public Const MAX_DAY As Long = 2958465       ' 31/12/9999
Public Const DATE1904_OFFSET As Long = 1462  ' chênh lệch hệ ngày 1904

Public Function IsValidDay(ByVal d As Double) As Boolean
    IsValidDay = (d = Int(d) And d >= MIN_DAY And d <= MAX_DAY)
End Function

' Ngày/tháng/năm -> DayNumber; 0 nếu không tồn tại (ví dụ 31/02).
Public Function DayFromCalendar(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Long
    Dim dt As Date
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    ' DateSerial tự "tràn" ngày không hợp lệ (31/02 -> 03/03) -> so lại.
    If Year(dt) <> y Or Month(dt) <> m Or Day(dt) <> d Then Exit Function
    If IsValidDay(CDbl(dt)) Then DayFromCalendar = CLng(CDbl(dt))
End Function

' DayNumber -> "dd/mm/yyyy".
Public Function FormatDay(ByVal dayNum As Long) As String
    Dim dt As Date
    dt = CDate(dayNum)
    FormatDay = Pad2(Day(dt)) & "/" & Pad2(Month(dt)) & "/" & CStr(Year(dt))
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    AllDigits = True
End Function

' Đọc ngày dạng chữ: dd/mm/yyyy (dấu / - .) hoặc yyyy-mm-dd; 0 nếu không đọc được.
Public Function ParseDateText(ByVal text As String) As Long
    Dim t As String, parts As Variant, i As Long, iso As Boolean
    t = Trim$(text)
    If Len(t) < 8 Or Len(t) > 10 Then Exit Function
    ' yyyy-mm-dd chỉ dùng dấu "-"; dd/mm/yyyy cho phép trộn "/", "-", "." (giống bản Office.js).
    iso = (Mid$(t, 5, 1) = "-")
    If Not iso Then t = Replace(Replace(t, "-", "/"), ".", "/")
    If iso Then parts = Split(t, "-") Else parts = Split(t, "/")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Not AllDigits(parts(i)) Then Exit Function
    Next i
    If iso Then
        If Len(parts(0)) <> 4 Or Len(parts(1)) > 2 Or Len(parts(2)) > 2 Then Exit Function
        ParseDateText = DayFromCalendar(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
    ElseIf Len(parts(0)) <= 2 And Len(parts(1)) <= 2 And Len(parts(2)) = 4 Then
        ParseDateText = DayFromCalendar(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
    End If
End Function

' Giá trị ô (Range.Value2) -> DayNumber; 0 nếu không phải ngày hợp lệ.
' Số: serial của Excel (bỏ phần giờ); chuỗi: ParseDateText.
Public Function ToDayNumber(ByVal v As Variant, Optional ByVal date1904 As Boolean = False) As Long
    Dim d As Double
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbDate
            d = Int(CDbl(v))
            If date1904 Then d = d + DATE1904_OFFSET
            If IsValidDay(d) Then ToDayNumber = CLng(d)
        Case vbString
            ToDayNumber = ParseDateText(v)
    End Select
End Function

' DayNumber -> serial theo hệ ngày của workbook để ghi vào ô.
Public Function ToWorkbookSerial(ByVal dayNum As Long, Optional ByVal date1904 As Boolean = False) As Long
    If date1904 Then ToWorkbookSerial = dayNum - DATE1904_OFFSET Else ToWorkbookSerial = dayNum
End Function

' Thứ trong tuần: "Thứ Hai" ... "Chủ nhật".
Public Function TenThu(ByVal dayNum As Long) As String
    Select Case Weekday(CDate(dayNum), vbSunday)
        Case 1: TenThu = "Chủ nhật"
        Case 2: TenThu = "Thứ Hai"
        Case 3: TenThu = "Thứ Ba"
        Case 4: TenThu = "Thứ Tư"
        Case 5: TenThu = "Thứ Năm"
        Case 6: TenThu = "Thứ Sáu"
        Case 7: TenThu = "Thứ Bảy"
    End Select
End Function

' Vị trí 3 từ khóa "ngày", "tháng", "năm" (không phân biệt hoa/thường, theo thứ tự); False nếu không phải mẫu ngày.
Private Function KeywordPositions(ByVal text As String, ByRef pd As Long, ByRef pm As Long, ByRef py As Long) As Boolean
    Dim lower As String
    lower = LCase$(text)
    If Len(lower) <> Len(text) Then Exit Function
    pd = InStr(1, lower, "ngày", vbBinaryCompare)
    If pd = 0 Then Exit Function
    pm = InStr(pd + Len("ngày"), lower, "tháng", vbBinaryCompare)
    If pm = 0 Then Exit Function
    py = InStr(pm + Len("tháng"), lower, "năm", vbBinaryCompare)
    KeywordPositions = (py > 0)
End Function

' Ô có phải mẫu "… Ngày … tháng … năm …" không (R2.15).
Public Function IsDatePattern(ByVal text As String) As Boolean
    Dim pd As Long, pm As Long, py As Long
    If Len(text) > 200 Then Exit Function
    IsDatePattern = KeywordPositions(text, pd, pm, py)
End Function

' Rải ngày vào mẫu chữ (R2.15): "Ngày ..... tháng ..... năm ....." -> "Ngày 15 tháng 06 năm 2026".
' Giữ phần chữ trước "Ngày" và cách viết hoa của từ khóa. Trả về chuỗi rỗng nếu không phải mẫu ngày.
Public Function FillDatePattern(ByVal text As String, ByVal dayNum As Long) As String
    Dim pd As Long, pm As Long, py As Long, dt As Date
    If Len(text) > 200 Then Exit Function
    If Not KeywordPositions(text, pd, pm, py) Then Exit Function
    dt = CDate(dayNum)
    FillDatePattern = Left$(text, pd - 1 + Len("ngày")) & " " & Pad2(Day(dt)) & " " & _
        Mid$(text, pm, Len("tháng")) & " " & Pad2(Month(dt)) & " " & _
        Mid$(text, py, Len("năm")) & " " & CStr(Year(dt))
End Function
