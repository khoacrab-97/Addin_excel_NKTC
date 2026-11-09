Attribute VB_Name = "modNKTC_Date"
'==============================================================================
' NHAT KY TC ? ngay thang (khong phu thuoc Excel).
' Ngay dung so nguyen theo he ngay 1900 cua Excel (DayNumber, kieu Long), vi du 46327 = 01/11/2026.
' Ngay cua VBA (Date) trung he 1900 cua Excel tu 01/03/1900 tro di.
'==============================================================================
Option Explicit

Public Const MIN_DAY As Long = 61            ' 01/03/1900
Public Const MAX_DAY As Long = 2958465       ' 31/12/9999
Public Const DATE1904_OFFSET As Long = 1462  ' chenh lech he ngay 1904

Public Function IsValidDay(ByVal d As Double) As Boolean
    IsValidDay = (d = Int(d) And d >= MIN_DAY And d <= MAX_DAY)
End Function

' Ngay/thang/nam -> DayNumber; 0 neu khong ton tai (vi du 31/02).
Public Function DayFromCalendar(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Long
    Dim dt As Date
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    ' DateSerial tu "tran" ngay khong hop le (31/02 -> 03/03) -> so lai.
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

' Doc ngay dang chu: dd/mm/yyyy (dau / - .) hoac yyyy-mm-dd; 0 neu khong doc duoc.
Public Function ParseDateText(ByVal text As String) As Long
    Dim t As String, parts As Variant, i As Long, iso As Boolean
    t = Trim$(text)
    If Len(t) < 8 Or Len(t) > 10 Then Exit Function
    ' yyyy-mm-dd chi dung dau "-"; dd/mm/yyyy cho phep tron "/", "-", "." (giong ban Office.js).
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

' Gia tri o (Range.Value2) -> DayNumber; 0 neu khong phai ngay hop le.
' So: serial cua Excel (bo phan gio); chuoi: ParseDateText.
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

' DayNumber -> serial theo he ngay cua workbook de ghi vao o.
Public Function ToWorkbookSerial(ByVal dayNum As Long, Optional ByVal date1904 As Boolean = False) As Long
    If date1904 Then ToWorkbookSerial = dayNum - DATE1904_OFFSET Else ToWorkbookSerial = dayNum
End Function

' Thu trong tuan: "Thu Hai" ... "Chu nhat".
Public Function TenThu(ByVal dayNum As Long) As String
    Select Case Weekday(CDate(dayNum), vbSunday)
        Case 1: TenThu = U("Ch\u1ee7 nh\u1eadt")
        Case 2: TenThu = U("Th\u1ee9 Hai")
        Case 3: TenThu = U("Th\u1ee9 Ba")
        Case 4: TenThu = U("Th\u1ee9 T\u01b0")
        Case 5: TenThu = U("Th\u1ee9 N\u0103m")
        Case 6: TenThu = U("Th\u1ee9 S\u00e1u")
        Case 7: TenThu = U("Th\u1ee9 B\u1ea3y")
    End Select
End Function

' Vi tri 3 tu khoa "ngay", "thang", "nam" (khong phan biet hoa/thuong, theo thu tu); False neu khong phai mau ngay.
Private Function KeywordPositions(ByVal text As String, ByRef pd As Long, ByRef pm As Long, ByRef py As Long) As Boolean
    Dim lower As String
    lower = LCase$(text)
    If Len(lower) <> Len(text) Then Exit Function
    pd = InStr(1, lower, U("ng\u00e0y"), vbBinaryCompare)
    If pd = 0 Then Exit Function
    pm = InStr(pd + Len(U("ng\u00e0y")), lower, U("th\u00e1ng"), vbBinaryCompare)
    If pm = 0 Then Exit Function
    py = InStr(pm + Len(U("th\u00e1ng")), lower, U("n\u0103m"), vbBinaryCompare)
    KeywordPositions = (py > 0)
End Function

' O co phai mau "? Ngay ? thang ? nam ?" khong (R2.15).
Public Function IsDatePattern(ByVal text As String) As Boolean
    Dim pd As Long, pm As Long, py As Long
    If Len(text) > 200 Then Exit Function
    IsDatePattern = KeywordPositions(text, pd, pm, py)
End Function

' Rai ngay vao mau chu (R2.15): "Ngay ..... thang ..... nam ....." -> "Ngay 15 thang 06 nam 2026".
' Giu phan chu truoc "Ngay" va cach viet hoa cua tu khoa. Tra ve chuoi rong neu khong phai mau ngay.
Public Function FillDatePattern(ByVal text As String, ByVal dayNum As Long) As String
    Dim pd As Long, pm As Long, py As Long, dt As Date
    If Len(text) > 200 Then Exit Function
    If Not KeywordPositions(text, pd, pm, py) Then Exit Function
    dt = CDate(dayNum)
    FillDatePattern = Left$(text, pd - 1 + Len(U("ng\u00e0y"))) & " " & Pad2(Day(dt)) & " " & _
        Mid$(text, pm, Len(U("th\u00e1ng"))) & " " & Pad2(Month(dt)) & " " & _
        Mid$(text, py, Len(U("n\u0103m"))) & " " & CStr(Year(dt))
End Function
