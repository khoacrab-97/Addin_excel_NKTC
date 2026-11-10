Attribute VB_Name = "modNKTC_Tests"
'==============================================================================
' NHAT KY TC ? kiem thu phan loi (tuong ung cac file *.test.ts cua ban Office.js).
' Chay: Application.Run "NhatKyThiCong_Core.xlam!modNKTC_Tests.NKTC_RunTests" -> "PASS n" hoac "FAIL ...".
'==============================================================================
Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As String
Private mLast As String

Private Sub Check(ByVal ok As Boolean, ByVal name As String)
    mLast = name
    If ok Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        mLog = mLog & vbLf & "  x " & name
    End If
End Sub

Private Sub Eq(ByVal actual As Variant, ByVal expected As Variant, ByVal name As String)
    Dim ok As Boolean
    If VarType(actual) = VarType(expected) Or (IsNumeric(actual) And IsNumeric(expected) And VarType(actual) <> vbString And VarType(expected) <> vbString) Then
        ok = (actual = expected)
    End If
    If Not ok Then name = name & U(" (\u0111\u01b0\u1ee3c: ") & CStr(actual) & U(" | c\u1ea7n: ") & CStr(expected) & ")"
    Check ok, name
End Sub

' Mang 2 chieu 1-based tu cac dong (moi dong la Array(...)).
Private Function Grid(ParamArray rows() As Variant) As Variant
    Dim r As Long, c As Long, nCols As Long, out() As Variant
    nCols = UBound(rows(0)) + 1
    ReDim out(1 To UBound(rows) + 1, 1 To nCols)
    For r = 0 To UBound(rows)
        For c = 0 To nCols - 1
            If c <= UBound(rows(r)) Then out(r + 1, c + 1) = rows(r)(c)
        Next c
    Next r
    Grid = out
End Function

Public Function NKTC_RunTests() As String
    Dim names As Variant, k As Long
    mPass = 0: mFail = 0: mLog = ""
    names = Array("TestText", "TestDates", "TestDatePattern", "TestColumns", "TestDailyLog", "TestSheetName", _
        "TestDays", "TestNavigation", "TestHeaderMatch", "TestWriteValue", "TestValidate", "TestDialogHelpers", "TestReference", "TestFormRows", "TestConfig", "TestUpdate", "TestPrint")
    For k = 0 To UBound(names)
        RunOne k, CStr(names(k))
    Next k
    If mFail = 0 Then
        NKTC_RunTests = "PASS " & mPass
    Else
        NKTC_RunTests = "FAIL " & mFail & "/" & (mPass + mFail) & mLog
    End If
    Debug.Print NKTC_RunTests
End Function

' Chay mot nhom kiem thu; loi khi chay duoc ghi vao ket qua (khong hien hop thoai) roi chay tiep nhom sau.
Private Sub RunOne(ByVal idx As Long, ByVal testName As String)
    On Error GoTo Loi
    mLast = U("(ch\u01b0a c\u00f3 ph\u00e9p ki\u1ec3m tra n\u00e0o)")
    Select Case idx
        Case 0: TestText
        Case 1: TestDates
        Case 2: TestDatePattern
        Case 3: TestColumns
        Case 4: TestDailyLog
        Case 5: TestSheetName
        Case 6: TestDays
        Case 7: TestNavigation
        Case 8: TestHeaderMatch
        Case 9: TestWriteValue
        Case 10: TestValidate
        Case 11: TestDialogHelpers
        Case 12: TestReference
        Case 13: TestFormRows
        Case 14: TestConfig
        Case 15: TestUpdate
        Case 16: TestPrint
    End Select
    Exit Sub
Loi:
    mFail = mFail + 1
    mLog = mLog & vbLf & "  x " & testName & U(": l\u1ed7i ") & Err.Number & U(" \u2013 ") & Err.Description & _
        U(" (sau ph\u00e9p ki\u1ec3m tra: ") & mLast & ")"
End Sub

Private Sub TestText()
    Eq U(U("Ng\u00e0y")), U("Ng\u00e0y"), U("U gi\u1ea3i m\u00e3")
    Eq RemoveDiacritics(U("Ng\u00e0y B\u1eaft \u0110\u1ea7u")), "ngay bat dau", U("b\u1ecf d\u1ea5u 1")
    Eq RemoveDiacritics(U("N\u1ed8I DUNG C\u00d4NG VI\u1ec6C")), "noi dung cong viec", U("b\u1ecf d\u1ea5u ch\u1eef hoa")
    Eq RemoveDiacritics(U("Th\u1eddi ti\u1ebft, \u01af\u1edbc l\u01b0\u1ee3ng, K\u1ef9 thu\u1eadt")), "thoi tiet, uoc luong, ky thuat", U("b\u1ecf d\u1ea5u 3")
    Eq NormalizeHeader(U("  Ng\u00e0y") & vbLf & U("b\u1eaft   \u0111\u1ea7u (*) ")), "ngay bat dau", U("chu\u1ea9n h\u00f3a ti\u00eau \u0111\u1ec1")
    Eq NormalizeHeader(U("Ban ch\u1ec9 huy - CBKT")), "ban chi huy cbkt", U("chu\u1ea9n h\u00f3a ti\u00eau \u0111\u1ec1 2")
    Eq NormalizeHeaderName(U("  TH\u1edcI   ti\u1ebft ")), NormalizeHeaderName(U("Th\u1eddi ti\u1ebft")), U("t\u00ean ti\u00eau \u0111\u1ec1 kh\u00f4ng ph\u00e2n bi\u1ec7t hoa/kho\u1ea3ng tr\u1eafng")
    Check NormalizeHeaderName("Thoi tiet") <> NormalizeHeaderName(U("Th\u1eddi ti\u1ebft")), U("t\u00ean ti\u00eau \u0111\u1ec1 gi\u1eef d\u1ea5u")
    Eq CellText(12), "12", U("CellText s\u1ed1")
    Check IsBlankValue(Empty) And IsBlankValue("  ") And Not IsBlankValue(0), "IsBlankValue"
End Sub

Private Sub TestDates()
    Const NOV1 As Long = 46327
    Eq DayFromCalendar(2026, 11, 1), NOV1, "DATE(2026,11,1)"
    Eq DayFromCalendar(1900, 3, 1), MIN_DAY, "MIN_DAY"
    Eq DayFromCalendar(9999, 12, 31), MAX_DAY, "MAX_DAY"
    Eq DayFromCalendar(2026, 2, 29), 0, U("29/02/2026 kh\u00f4ng t\u1ed3n t\u1ea1i")
    Eq DayFromCalendar(2026, 4, 31), 0, U("31/04 kh\u00f4ng t\u1ed3n t\u1ea1i")
    Eq DayFromCalendar(1900, 2, 28), 0, U("tr\u01b0\u1edbc MIN_DAY")
    Check DayFromCalendar(2028, 2, 29) > 0, "29/02/2028"
    Eq FormatDay(NOV1), "01/11/2026", "FormatDay"
    Eq FormatDay(NOV1 + 60), "31/12/2026", "FormatDay +60"
    Dim ok As Variant
    For Each ok In Array("01/11/2026", "1/11/2026", "01-11-2026", "01.11.2026", "  01/11/2026  ", "2026-11-01", "01/11-2026")
        Eq ParseDateText(ok), NOV1, U("\u0111\u1ecdc ") & ok
    Next ok
    For Each ok In Array("31/02/2026", "11/2026", "01/11/26", "abc", "", "01/11/2026 08:00")
        Eq ParseDateText(ok), 0, U("t\u1eeb ch\u1ed1i ") & ok
    Next ok
    Eq ToDayNumber(CDbl(NOV1) + 0.75), NOV1, U("b\u1ecf ph\u1ea7n gi\u1edd")
    Eq ToDayNumber("01/11/2026"), NOV1, U("ToDayNumber ch\u1eef")
    Eq ToDayNumber(0), 0, U("0 kh\u00f4ng ph\u1ea3i ng\u00e0y")
    Eq ToDayNumber(-5), 0, U("s\u1ed1 \u00e2m")
    Eq ToDayNumber(CDbl(MAX_DAY) + 1), 0, U("qu\u00e1 MAX_DAY")
    Eq ToDayNumber(True), 0, "logic"
    Eq ToDayNumber(Empty), 0, U("tr\u1ed1ng")
    Eq ToDayNumber(NOV1 - 1462, True), NOV1, U("h\u1ec7 1904")
    Eq ToWorkbookSerial(NOV1, True), NOV1 - 1462, U("ghi h\u1ec7 1904")
End Sub

Private Sub TestDatePattern()
    Const D As Long = 46188 ' 15/06/2026, thu Hai
    Eq FillDatePattern(U("Ng\u00e0y ..... th\u00e1ng ..... n\u0103m .........."), D), U("Ng\u00e0y 15 th\u00e1ng 06 n\u0103m 2026"), U("r\u1ea3i ng\u00e0y")
    Eq FillDatePattern(U("Ng\u00e0y 01 th\u00e1ng 01 n\u0103m 2025"), D + 1), U("Ng\u00e0y 16 th\u00e1ng 06 n\u0103m 2026"), U("\u0111i\u1ec1n l\u1ea1i")
    Eq FillDatePattern(U("H\u00e0 N\u1ed9i, ng\u00e0y \u2026 th\u00e1ng \u2026 n\u0103m \u2026"), D), U("H\u00e0 N\u1ed9i, ng\u00e0y 15 th\u00e1ng 06 n\u0103m 2026"), U("gi\u1eef ph\u1ea7n tr\u01b0\u1edbc")
    Eq FillDatePattern(U("NG\u00c0Y __ TH\u00c1NG __ N\u0102M ____"), D), U("NG\u00c0Y 15 TH\u00c1NG 06 N\u0102M 2026"), U("gi\u1eef ch\u1eef hoa")
    Eq FillDatePattern(U("N\u1ed9i dung c\u00f4ng vi\u1ec7c"), D), "", U("kh\u00f4ng ph\u1ea3i m\u1eabu")
    Eq FillDatePattern(U("n\u0103m th\u00e1ng ng\u00e0y"), D), "", U("sai th\u1ee9 t\u1ef1")
    Check IsDatePattern(U("Ng\u00e0y ... th\u00e1ng ... n\u0103m ...")) And Not IsDatePattern(U("Th\u1eddi ti\u1ebft")), "IsDatePattern"
    Eq TenThu(D), U("Th\u1ee9 Hai"), U("th\u1ee9 Hai")
    Eq TenThu(D + 6), U("Ch\u1ee7 nh\u1eadt"), U("Ch\u1ee7 nh\u1eadt")
End Sub

Private Sub TestColumns()
    Dim c As Long, s As Long, e As Long
    DetectColumns Array("STT", U("N\u1ed9i dung c\u00f4ng vi\u1ec7c"), U("Ng\u00e0y b\u1eaft \u0111\u1ea7u"), U("Ng\u00e0y k\u1ebft th\u00fac")), c, s, e
    Check c = 1 And s = 2 And e = 3, U("ti\u00eau \u0111\u1ec1 chu\u1ea9n")
    DetectColumns Array(U("T\u00caN C\u00d4NG VI\u1ec6C"), "Tu ngay", "Den ngay"), c, s, e
    Check c = 0 And s = 1 And e = 2, U("kh\u00f4ng d\u1ea5u, vi\u1ebft hoa")
    DetectColumns Array(U("H\u1ea1ng m\u1ee5c c\u00f4ng vi\u1ec7c"), U("Th\u1eddi gian b\u1eaft \u0111\u1ea7u"), U("Th\u1eddi gian k\u1ebft th\u00fac")), c, s, e
    Check c = 0 And s = 1 And e = 2, U("ch\u1ee9a c\u1ee5m t\u1eeb")
    DetectColumns Array(U("T\u00caN H\u1ea0NG M\u1ee4C/ C\u00d4NG T\u00c1C"), U("B\u1eaeT \u0110\u1ea6U"), U("K\u1ebeT TH\u00daC")), c, s, e
    Check c = 0 And s = 1 And e = 2, U("h\u1ea1ng m\u1ee5c / c\u00f4ng t\u00e1c")
    DetectColumns Array("STT", U("H\u1ea1ng m\u1ee5c"), U("T\u1eeb ng\u00e0y"), U("\u0110\u1ebfn ng\u00e0y")), c, s, e
    Check c = 1 And s = 2 And e = 3, U("t\u1eeb ng\u00e0y / \u0111\u1ebfn ng\u00e0y")
    DetectColumns Array(U("N\u1ed9i dung nghi\u1ec7m thu"), U("N\u1ed9i dung c\u00f4ng vi\u1ec7c"), U("Ng\u00e0y b\u1eaft \u0111\u1ea7u"), U("Ng\u00e0y k\u1ebft th\u00fac")), c, s, e
    Check c = 1 And s = 2 And e = 3, U("kh\u00f4ng nh\u1ea7m c\u1ed9t n\u1ed9i dung ph\u1ee5")
    DetectColumns Array(1, Empty, U("Ghi ch\u00fa")), c, s, e
    Check c = -1 And s = -1 And e = -1, U("kh\u00f4ng nh\u1eadn di\u1ec7n \u0111\u01b0\u1ee3c")
    Check IsCompleteColumns(0, 1, 2) And Not IsCompleteColumns(0, 1, -1) And Not IsCompleteColumns(0, 1, 1), "IsCompleteColumns"
End Sub

Private Sub TestDailyLog()
    Dim items() As NKTC_WorkItem, n As Long, issues() As NKTC_RowIssue, ni As Long, blank As Long
    Dim entries() As NKTC_DailyEntry, ne As Long, f As Long, l As Long, msg As String
    Dim v As Variant
    Const D As Long = 46188
    v = Grid(Array(U("N\u1ed9i dung"), U("B\u1eaft \u0111\u1ea7u"), U("K\u1ebft th\u00fac")), _
             Array(U("- \u0110\u00e0o \u0111\u1ea5t"), CDbl(D), CDbl(D + 2)), _
             Array(U("\u0110\u1ed5 b\u00ea t\u00f4ng"), "16/06/2026", Empty), _
             Array(Empty, Empty, Empty), _
             Array("", CDbl(D), CDbl(D)), _
             Array(U("Sai ng\u00e0y"), "abc", Empty), _
             Array(U("Ng\u01b0\u1ee3c"), CDbl(D + 3), CDbl(D)))
    ParseWorkItems v, 0, 1, 2, 1, False, items, n, issues, ni, blank
    Eq n, 2, U("s\u1ed1 c\u00f4ng vi\u1ec7c h\u1ee3p l\u1ec7")
    Eq ni, 3, U("s\u1ed1 d\u00f2ng l\u1ed7i")
    Eq blank, 1, U("d\u00f2ng tr\u1ed1ng")
    Eq items(1).EndDay, D + 1, U("thi\u1ebfu ng\u00e0y k\u1ebft th\u00fac -> 1 ng\u00e0y")
    Eq issues(0).RowNumber, 5, U("s\u1ed1 d\u00f2ng l\u1ed7i \u0111\u1ea7u")
    Eq issues(0).Reason, U("Thi\u1ebfu n\u1ed9i dung c\u00f4ng vi\u1ec7c."), U("l\u00fd do thi\u1ebfu n\u1ed9i dung")
    Eq issues(1).Reason, U("Ng\u00e0y b\u1eaft \u0111\u1ea7u kh\u00f4ng h\u1ee3p l\u1ec7."), U("l\u00fd do ng\u00e0y sai")
    Eq issues(2).Reason, U("Ng\u00e0y k\u1ebft th\u00fac tr\u01b0\u1edbc ng\u00e0y b\u1eaft \u0111\u1ea7u."), U("l\u00fd do ng\u01b0\u1ee3c ng\u00e0y")
    msg = BuildDailyLog(items, n, entries, ne, f, l)
    Eq msg, "", U("BuildDailyLog th\u00e0nh c\u00f4ng")
    Eq ne, 3, U("s\u1ed1 ng\u00e0y")
    Eq entries(1).Content, U("- \u0110\u00e0o \u0111\u1ea5t") & vbLf & U("\u0110\u1ed5 b\u00ea t\u00f4ng"), U("g\u1ed9p n\u1ed9i dung, gi\u1eef nguy\u00ean d\u1ea5u -")
    Eq entries(1).ItemCount, 2, U("\u0111\u1ebfm c\u00f4ng vi\u1ec7c")
    items(0).EndDay = D + MAX_SPAN_DAYS
    msg = BuildDailyLog(items, n, entries, ne, f, l)
    Check InStr(msg, U("v\u01b0\u1ee3t gi\u1edbi h\u1ea1n")) > 0, U("ch\u1eb7n kho\u1ea3ng qu\u00e1 d\u00e0i")
    msg = BuildDailyLog(items, 0, entries, ne, f, l)
    Eq msg, U("Kh\u00f4ng c\u00f3 c\u00f4ng vi\u1ec7c h\u1ee3p l\u1ec7 \u0111\u1ec3 xu\u1ea5t."), U("kh\u00f4ng c\u00f3 c\u00f4ng vi\u1ec7c")
End Sub

Private Sub TestSheetName()
    Dim names As New Collection
    Eq UniqueSheetName(names), "NK_TheoNgay", U("t\u00ean g\u1ed1c")
    names.Add "nk_theongay": names.Add "NK_TheoNgay_2"
    Eq UniqueSheetName(names), "NK_TheoNgay_3", U("t\u00ean ti\u1ebfp theo, kh\u00f4ng ph\u00e2n bi\u1ec7t hoa/th\u01b0\u1eddng")
End Sub

Private Sub TestDays()
    Dim v As Variant, days() As NKTC_SourceDay, n As Long, issues As Collection
    Const D As Long = 46188
    v = Grid(Array(U("Ng\u00e0y"), U("Th\u1eddi ti\u1ebft")), Array(CDbl(D + 1), "a"), Array(CDbl(D), "b"), Array(Empty, "c"), _
             Array("xx", "d"), Array(CDbl(D), "e"))
    Eq DetectDateColumn(v), 0, U("c\u1ed9t Ng\u00e0y theo ti\u00eau \u0111\u1ec1")
    BuildDayIndex v, 0, 1, False, days, n, issues
    Eq n, 2, U("s\u1ed1 ng\u00e0y")
    Eq days(0).DayNum, D, U("s\u1eafp x\u1ebfp t\u0103ng d\u1ea7n")
    Eq days(0).RowNumber, 3, U("s\u1ed1 d\u00f2ng")
    Eq issues.Count, 2, U("b\u00e1o ng\u00e0y sai v\u00e0 tr\u00f9ng")
    Check InStr(issues(2), U("Tr\u00f9ng ng\u00e0y 15/06/2026 v\u1edbi d\u00f2ng 3")) > 0, U("n\u1ed9i dung b\u00e1o tr\u00f9ng")
    Eq IndexOfDay(days, n, D + 1), 1, "IndexOfDay"
    v = Grid(Array("A", U("Ghi ch\u00fa ng\u00e0y")), Array("x", CDbl(D)))
    Eq DetectDateColumn(v), 1, U("ti\u00eau \u0111\u1ec1 c\u00f3 t\u1eeb ngay")
    v = Grid(Array("A", "B"), Array("x", CDbl(D)), Array("y", "16/06/2026"))
    Eq DetectDateColumn(v), 1, U("c\u1ed9t to\u00e0n ng\u00e0y")
End Sub

Private Sub TestNavigation()
    Dim v As Variant, days() As NKTC_SourceDay, n As Long, issues As Collection, r As NKTC_DayResolution
    Const D As Long = 46188
    v = Grid(Array(U("Ng\u00e0y")), Array(CDbl(D)), Array(CDbl(D + 1)), Array(CDbl(D + 2)))
    BuildDayIndex v, 0, 1, False, days, n, issues
    r = ResolveDay(days, n, 0, CMD_FIRST): Check r.Ok And r.TargetIndex = 0, U("ng\u00e0y \u0111\u1ea7u")
    r = ResolveDay(days, n, 0, CMD_LAST): Check r.Ok And r.TargetIndex = 2, U("ng\u00e0y cu\u1ed1i")
    r = ResolveDay(days, n, D, CMD_NEXT): Check r.Ok And r.TargetIndex = 1, U("ng\u00e0y sau")
    r = ResolveDay(days, n, D + 2, CMD_NEXT): Check (Not r.Ok) And r.IsInfo And InStr(r.Message, U("\u0110ang \u1edf ng\u00e0y cu\u1ed1i c\u00f9ng (17/06/2026)")) > 0, U("\u0111\u00e3 \u1edf cu\u1ed1i")
    r = ResolveDay(days, n, D, CMD_PREV): Check (Not r.Ok) And r.IsInfo And InStr(r.Message, U("\u0110ang \u1edf ng\u00e0y \u0111\u1ea7u ti\u00ean")) > 0, U("\u0111\u00e3 \u1edf \u0111\u1ea7u")
    r = ResolveDay(days, n, 0, CMD_NEXT): Check r.Ok And r.TargetIndex = 0 And r.Notice = "", U("ch\u01b0a hi\u1ec3n th\u1ecb -> ng\u00e0y \u0111\u1ea7u")
    r = ResolveDay(days, n, D + 9, CMD_PREV): Check r.Ok And r.TargetIndex = 0 And InStr(r.Notice, U("Kh\u00f4ng c\u00f2n ng\u00e0y")) > 0, U("ng\u00e0y c\u0169 kh\u00f4ng c\u00f2n")
    r = ResolveDay(days, 0, 0, CMD_FIRST): Check (Not r.Ok) And (Not r.IsInfo), U("kh\u00f4ng c\u00f3 ng\u00e0y")
    r = ResolveDay(days, n, 0, CMD_GOTO, D + 1): Check r.Ok And r.TargetIndex = 1, U("\u0111i t\u1edbi ng\u00e0y c\u00f3")
    r = ResolveDay(days, n, D, CMD_GOTO, D + 9)
    Check (Not r.Ok) And (Not r.IsInfo) And InStr(r.Message, U("Kh\u00f4ng c\u00f3 ng\u00e0y 24/06/2026 trong v\u00f9ng ngu\u1ed3n")) > 0, U("\u0111i t\u1edbi ng\u00e0y kh\u00f4ng c\u00f3")
End Sub

Private Function Strs(ParamArray items() As Variant) As String()
    Dim out() As String, i As Long
    ReDim out(0 To UBound(items))
    For i = 0 To UBound(items)
        out(i) = items(i)
    Next i
    Strs = out
End Function

Private Sub TestHeaderMatch()
    Dim oldH() As String, newH() As String, m() As Long
    oldH = Strs(U("Ng\u00e0y"), U("N\u1ed9i dung"), U("Th\u1eddi ti\u1ebft"), U("Ghi ch\u00fa"), U("Ghi ch\u00fa"))
    newH = Strs(U("Ng\u00e0y"), U("N\u1ed9i dung"), U("S\u1ed1 c\u00f4ng nh\u00e2n"), U("Th\u1eddi ti\u1ebft"), U("Ghi ch\u00fa"), U("Ghi ch\u00fa"))
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = 1 And m(2) = 3 And m(3) = 4 And m(4) = 5, U("ch\u00e8n c\u1ed9t gi\u1eefa")
    oldH = Strs(U("Ghi ch\u00fa"), "A", U("Ghi ch\u00fa"))
    newH = Strs(U("Ghi ch\u00fa"), U("Ghi ch\u00fa"), "A")
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = 2 And m(2) = 1, U("tr\u00f9ng t\u00ean theo th\u1ee9 t\u1ef1")
    oldH = Strs(U("Ng\u00e0y"), "", U("Th\u1eddi ti\u1ebft"))
    newH = Strs(U("ng\u00e0y "), "", U("Th\u1eddi ti\u1ebft (s\u00e1ng)"))
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = -1 And m(2) = -1, U("kh\u00f4ng ph\u00e2n bi\u1ec7t hoa/kho\u1ea3ng tr\u1eafng; tr\u1ed1ng v\u00e0 \u0111\u1ed5i t\u00ean kh\u00f4ng kh\u1edbp")
End Sub

Private Sub TestWriteValue()
    Const D As Long = 46188
    Eq ToCellInput("=SUM(A1)"), "'=SUM(A1)", U("chu\u1ed7i th\u00eam d\u1ea5u '")
    Eq ToCellInput(""), "", U("chu\u1ed7i r\u1ed7ng")
    Eq ToCellInput(5), 5, U("s\u1ed1 gi\u1eef nguy\u00ean")
    Eq ToCellInput(Empty), "", U("tr\u1ed1ng")
    Eq WriteValue(CDbl(D), True, U("Ng\u00e0y ..... th\u00e1ng ..... n\u0103m .........."), False), U("'Ng\u00e0y 15 th\u00e1ng 06 n\u0103m 2026"), U("c\u1ed9t Ng\u00e0y v\u00e0o \u00f4 m\u1eabu ng\u00e0y")
    Eq WriteValue(CDbl(D), True, "", False), CDbl(D), U("c\u1ed9t Ng\u00e0y v\u00e0o \u00f4 tr\u1ed1ng -> s\u1ed1")
    Eq WriteValue(U("N\u1eafng"), False, U("Ng\u00e0y ... th\u00e1ng ... n\u0103m ..."), False), U("'N\u1eafng"), U("c\u1ed9t kh\u00e1c gi\u1eef nguy\u00ean")
End Sub

Private Sub TestValidate()
    Dim maps(0 To 2) As NKTC_Mapping, errs As Collection
    maps(0).Column = 0: maps(0).SheetName = U("M\u1eabu NKTC"): maps(0).Address = "E2"
    maps(1).Column = 1: maps(1).SheetName = U("m\u1eabu nktc"): maps(1).Address = "e2"
    maps(2).Column = 9: maps(2).SheetName = U("M\u1eabu NKTC"): maps(2).Address = "C6"
    Set errs = ValidateMappings(maps, 3, 5)
    Eq errs.Count, 2, U("\u00f4 tr\u00f9ng + c\u1ed9t ngo\u00e0i v\u00f9ng")
    Eq MappingKey(maps, 2), U("m\u1eabu nktc!E2|m\u1eabu nktc!E2"), "MappingKey"
    Set errs = ValidateMappings(maps, 0, 5)
    Eq errs.Count, 1, U("ch\u01b0a g\u00e1n")
End Sub

Private Sub TestDialogHelpers()
    Dim cc As Long, cs As Long, ce As Long
    Eq ColumnLetter(0), "A", U("c\u1ed9t A")
    Eq ColumnLetter(25), "Z", U("c\u1ed9t Z")
    Eq ColumnLetter(26), "AA", U("c\u1ed9t AA")
    Eq ColumnLetter(16383), "XFD", U("c\u1ed9t cu\u1ed1i XFD")
    Eq ColumnOptionLabel(U("  N\u1ed9i dung") & vbLf & U("c\u00f4ng vi\u1ec7c "), 2), U("C\u1ed9t C \u2013 N\u1ed9i dung c\u00f4ng vi\u1ec7c"), U("nh\u00e3n c\u1ed9t")
    Eq ColumnOptionLabel(Empty, 0), U("C\u1ed9t A \u2013 (kh\u00f4ng c\u00f3 ti\u00eau \u0111\u1ec1)"), U("nh\u00e3n c\u1ed9t tr\u1ed1ng")
    Eq ColumnOptionLabel(2026, 1), U("C\u1ed9t B \u2013 2026"), U("nh\u00e3n c\u1ed9t s\u1ed1")
    Eq Len(ShortLabel(String(80, "x"))), 60, U("nh\u00e3n d\u00e0i c\u1eaft c\u00f2n 60")
    Eq ParseRowCount(" 10 "), 10, U("s\u1ed1 d\u00f2ng 10")
    Eq ParseRowCount("1000"), 1000, U("s\u1ed1 d\u00f2ng t\u1ed1i \u0111a")
    Eq ParseRowCount("1001"), 0, U("qu\u00e1 s\u1ed1 d\u00f2ng t\u1ed1i \u0111a")
    Eq ParseRowCount("0"), 0, U("s\u1ed1 d\u00f2ng 0")
    Eq ParseRowCount("5a"), 0, U("s\u1ed1 d\u00f2ng c\u00f3 ch\u1eef")
    Eq ParseRowCount(""), 0, U("s\u1ed1 d\u00f2ng tr\u1ed1ng")
    ' Bang do Tao bang sinh ra phai duoc Xuat theo ngay nhan dien du 3 cot (R1.12).
    DetectColumns InputTableHeaders(), cc, cs, ce
    Check cc = 1 And cs = 2 And ce = 3, U("T\u1ea1o b\u1ea3ng -> nh\u1eadn di\u1ec7n \u0111\u1ee7 c\u1ed9t")
End Sub

Private Sub TestReference()
    Dim sh As String, ad As String
    Eq NormalizeCellAddress(" $e$2 "), "E2", U("\u0111\u1ecba ch\u1ec9 \u00f4 b\u1ecf $, vi\u1ebft hoa")
    Eq NormalizeCellAddress("XFD1048576"), "XFD1048576", U("\u00f4 cu\u1ed1i sheet")
    Eq NormalizeCellAddress("XFE1"), "", U("c\u1ed9t v\u01b0\u1ee3t XFD")
    Eq NormalizeCellAddress("A0"), "", U("d\u00f2ng 0")
    Eq NormalizeCellAddress("A1:B2"), "", U("v\u00f9ng kh\u00f4ng ph\u1ea3i m\u1ed9t \u00f4")
    Check ParseCellReference(U("'M\u1eabu NKTC'!$E$2"), sh, ad) And sh = U("M\u1eabu NKTC") And ad = "E2", U("\u00f4 c\u00f3 t\u00ean sheet trong nh\u00e1y")
    Check ParseCellReference(U("M\u1eabu NKTC!e2"), sh, ad) And sh = U("M\u1eabu NKTC") And ad = "E2", U("\u00f4 c\u00f3 t\u00ean sheet kh\u00f4ng nh\u00e1y")
    Check ParseCellReference("E2", sh, ad) And sh = "" And ad = "E2", U("\u00f4 kh\u00f4ng ghi sheet")
    Check ParseCellReference("'It''s'!B3", sh, ad) And sh = "It's", U("t\u00ean sheet c\u00f3 d\u1ea5u nh\u00e1y")
    Check Not ParseCellReference("A1:B2", sh, ad), U("v\u00f9ng nhi\u1ec1u \u00f4 b\u1ecb t\u1eeb ch\u1ed1i")
    Check Not ParseCellReference("Sh[1]!A1", sh, ad), U("t\u00ean sheet c\u00f3 k\u00fd t\u1ef1 c\u1ea5m")
    Check Not ParseCellReference("", sh, ad), U("\u00f4 tr\u1ed1ng")
    Eq FormatReference(U("M\u1eabu NKTC"), "E2"), U("M\u1eabu NKTC!E2"), U("hi\u1ec3n th\u1ecb \u00f4")
    Eq QuotedReference("It's", "B3"), "'It''s'!B3", U("tham chi\u1ebfu cho Excel")
End Sub

Private Sub TestFormRows()
    Dim cols() As Long, targets() As String, maps() As NKTC_Mapping, n As Long, errors As Collection, h() As String
    ReDim cols(0 To 3): ReDim targets(0 To 3)
    cols(0) = 0: targets(0) = U("M\u1eabu NKTC!E2")
    cols(1) = -1: targets(1) = ""
    cols(2) = 1: targets(2) = "e2"
    cols(3) = 2: targets(3) = "x"
    Set errors = BuildMappingsFromRows(cols, targets, 4, 3, U("M\u1eabu NKTC"), maps, n)
    Eq errors.Count, 2, U("s\u1ed1 l\u1ed7i d\u00f2ng g\u00e1n")
    Eq errors(1), U("D\u00f2ng 3: \u00f4 M\u1eabu NKTC!E2 \u0111\u00e3 \u0111\u01b0\u1ee3c g\u00e1n \u1edf d\u00f2ng 1."), U("tr\u00f9ng \u00f4 (\u00f4 kh\u00f4ng ghi sheet thu\u1ed9c sheet m\u1eb7c \u0111\u1ecbnh)")
    Eq errors(2), U("D\u00f2ng 4: ""x"" kh\u00f4ng ph\u1ea3i m\u1ed9t \u00f4 h\u1ee3p l\u1ec7 (v\u00ed d\u1ee5 M\u1eabu NKTC!E2)."), U("\u00f4 kh\u00f4ng h\u1ee3p l\u1ec7")
    Eq n, 0, U("c\u00f3 l\u1ed7i th\u00ec kh\u00f4ng c\u00f3 g\u00e1n")
    cols(2) = -1: targets(2) = "F3"
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, U("M\u1eabu NKTC"), maps, n)
    Eq errors(1), U("D\u00f2ng 3: ch\u01b0a ch\u1ecdn c\u1ed9t ngu\u1ed3n."), U("thi\u1ebfu c\u1ed9t")
    cols(2) = 1: targets(2) = ""
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, U("M\u1eabu NKTC"), maps, n)
    Eq errors(1), U("D\u00f2ng 3: ch\u01b0a ch\u1ecdn \u00f4 tr\u00ean form."), U("thi\u1ebfu \u00f4")
    targets(2) = "F3"
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, "Form", maps, n)
    Eq errors.Count, 0, U("d\u00f2ng g\u00e1n h\u1ee3p l\u1ec7")
    Check n = 2 And maps(1).Column = 1 And maps(1).SheetName = "Form" And maps(1).Address = "F3", U("g\u00e1n d\u00f9ng sheet m\u1eb7c \u0111\u1ecbnh")
    Set errors = BuildMappingsFromRows(cols, targets, 0, 3, "Form", maps, n)
    Check errors.Count = 1 And InStr(errors(1), U("Ch\u01b0a c\u00f3 d\u00f2ng g\u00e1n")) > 0, U("ch\u01b0a c\u00f3 d\u00f2ng g\u00e1n")
    h = Strs("a", " ", "b", "")
    Eq BlankHeaderList(h, 4), "2, 4", U("ti\u00eau \u0111\u1ec1 tr\u1ed1ng")
    h = Strs("a", "b")
    Eq BlankHeaderList(h, 2), "", U("kh\u00f4ng c\u00f3 ti\u00eau \u0111\u1ec1 tr\u1ed1ng")
End Sub

Private Sub TestConfig()
    Dim cfg As NKTC_Config, back As NKTC_Config, maps() As NKTC_Mapping, s As String, msg As String, h() As String
    s = "a\b" & vbTab & "c" & vbLf & "d" & vbCr & "e"
    Eq UnescapeField(EscapeField(s)), s, U("m\u00e3 h\u00f3a tr\u01b0\u1eddng gi\u1eef nguy\u00ean n\u1ed9i dung")
    Check InStr(EscapeField(s), vbTab) = 0 And InStr(EscapeField(s), vbLf) = 0, U("tr\u01b0\u1eddng m\u00e3 h\u00f3a kh\u00f4ng ch\u1ee9a Tab/xu\u1ed1ng d\u00f2ng")
    Eq XmlText("a<b&c>" & vbTab & vbLf), "a&lt;b&amp;c&gt;&#9;&#10;", U("ch\u1eef trong XML")

    h = Strs(U("Ng\u00e0y"), U("Th\u1eddi ti\u1ebft") & vbLf & U("(s\u00e1ng)"), U("N\u1ed9i dung"))
    SetHeaders cfg, h, 3
    ReDim maps(0 To 1)
    maps(0).Column = 1: maps(0).SheetName = U("M\u1eabu NKTC"): maps(0).Address = "E5"
    maps(1).Column = 2: maps(1).SheetName = U("M\u1eabu NKTC"): maps(1).Address = "B9"
    SetMaps cfg, maps, 2
    cfg.SourceSheet = "NK_TheoNgay": cfg.SourceAddress = "A1:C40"
    cfg.DateColumn = 0: cfg.CurrentDay = 46188: cfg.ConfirmedKey = MappingKey(maps, 2)
    Check ParseConfig(SerializeConfig(cfg), back), U("\u0111\u1ecdc l\u1ea1i c\u1ea5u h\u00ecnh \u0111\u00e3 l\u01b0u")
    Check back.SourceSheet = "NK_TheoNgay" And back.SourceAddress = "A1:C40" And back.DateColumn = 0 And _
          back.CurrentDay = 46188 And back.ConfirmedKey = cfg.ConfirmedKey, U("c\u1ea5u h\u00ecnh: v\u00f9ng ngu\u1ed3n, ng\u00e0y, x\u00e1c nh\u1eadn")
    Check back.HeaderCount = 3 And back.Headers(1) = U("Th\u1eddi ti\u1ebft") & vbLf & U("(s\u00e1ng)"), U("c\u1ea5u h\u00ecnh: ti\u00eau \u0111\u1ec1 nhi\u1ec1u d\u00f2ng")
    Check back.MapCount = 2 And back.Maps(1).Column = 2 And back.Maps(1).SheetName = U("M\u1eabu NKTC") And _
          back.Maps(1).Address = "B9", U("c\u1ea5u h\u00ecnh: d\u00f2ng g\u00e1n")
    Check Not ParseConfig("abc", back), U("c\u1ea5u h\u00ecnh h\u1ecfng b\u1ecb b\u1ecf qua")
    Check Not ParseConfig(Replace(SerializeConfig(cfg), "date" & vbTab & "0", "date" & vbTab & "7"), back), _
        U("c\u1ed9t Ng\u00e0y ngo\u00e0i v\u00f9ng b\u1ecb t\u1eeb ch\u1ed1i")

    ' Khop lai theo ten tieu de: chen cot giua -> cot da gan dich theo.
    h = Strs(U("Ng\u00e0y"), U("Th\u1eddi ti\u1ebft"), U("N\u1ed9i dung"))
    SetHeaders cfg, h, 3
    h = Strs(U("Ng\u00e0y"), U("S\u1ed1 c\u00f4ng nh\u00e2n"), U("TH\u1edcI  TI\u1ebeT"), U("N\u1ed9i dung"))
    msg = RemapConfig(cfg, h, 4, "NK!A1:D9")
    Eq msg, "", U("kh\u1edbp l\u1ea1i sau khi ch\u00e8n c\u1ed9t")
    Check cfg.DateColumn = 0 And cfg.Maps(0).Column = 2 And cfg.Maps(1).Column = 3 And cfg.HeaderCount = 4, _
        U("v\u1ecb tr\u00ed c\u1ed9t sau khi ch\u00e8n c\u1ed9t")
    h = Strs(U("Ng\u00e0y"), U("N\u1ed9i dung"))
    msg = RemapConfig(cfg, h, 2, "NK!A1:B9")
    Check InStr(msg, U("Kh\u00f4ng t\u00ecm th\u1ea5y c\u1ed9t ""TH\u1edcI  TI\u1ebeT"" (g\u00e1n v\u00e0o M\u1eabu NKTC!E5)")) > 0, U("b\u00e1o c\u1ed9t b\u1ecb x\u00f3a k\u00e8m \u00f4 \u0111\u00edch")
    Check cfg.Maps(0).Column = 2, U("c\u1ed9t thi\u1ebfu th\u00ec kh\u00f4ng \u0111\u1ed5i c\u1ea5u h\u00ecnh")
End Sub

Private Sub TestUpdate()
    Dim ver As String, f As String, sha As String, size As Long, notes As String, b() As Byte
    ParseManifest "version=3.1.0" & vbCrLf & "file=NhatKyThiCong_Core.xlam" & vbCrLf & "sha256=" & String(64, "A") & vbLf & _
        "size=12345" & vbLf & U("notes=S\u1eeda l\u1ed7i") & vbLf, ver, f, sha, size, notes
    Check ver = "3.1.0" And f = "NhatKyThiCong_Core.xlam" And sha = String(64, "a") And size = 12345 And notes = U("S\u1eeda l\u1ed7i"), _
        U("\u0111\u1ecdc version.txt")
    ParseManifest "version=3.1" & vbLf & "file=..\x.xlam" & vbLf & "sha256=xyz", ver, f, sha, size, notes
    Check ver = "3.1" And f = "" And sha = "", U("version.txt: t\u00ean t\u1ec7p c\u00f3 \u0111\u01b0\u1eddng d\u1eabn, SHA-256 sai b\u1ecb b\u1ecf")
    ParseManifest "file=https://x/y.xlam" & vbLf & "version=3.0.0-beta", ver, f, sha, size, notes
    Check ver = "" And f = "", U("version.txt: phi\u00ean b\u1ea3n sai, t\u00ean t\u1ec7p l\u00e0 URL b\u1ecb b\u1ecf")
    Eq CompareVersions("3.10.0", "3.9.0"), 1, U("so phi\u00ean b\u1ea3n 3.10 > 3.9")
    Eq CompareVersions("3.0.0", "3.0.0"), 0, U("so phi\u00ean b\u1ea3n b\u1eb1ng")
    Eq CompareVersions("3.0", "3.0.1"), -1, U("so phi\u00ean b\u1ea3n thi\u1ebfu ph\u1ea7n")
    Check IsVersion("3.0.0") And Not IsVersion("3") And Not IsVersion("3.a.0") And Not IsVersion("3..0"), U("ki\u1ec3m tra d\u1ea1ng phi\u00ean b\u1ea3n")
    Check IsAllowedUpdateUrl("https://khoacrab-97.github.io/x/") And IsAllowedUpdateUrl("http://127.0.0.1:8080/"), U("\u0111\u1ecba ch\u1ec9 c\u1eadp nh\u1eadt h\u1ee3p l\u1ec7")
    Check Not IsAllowedUpdateUrl("http://example.com/") And Not IsAllowedUpdateUrl("https://u:p@example.com/") And _
          Not IsAllowedUpdateUrl("http://localhost.evil.com/"), U("\u0111\u1ecba ch\u1ec9 c\u1eadp nh\u1eadt b\u1ecb t\u1eeb ch\u1ed1i")
    b = StrConv("abc", vbFromUnicode)
    Eq Sha256Hex(b), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", U("SHA-256 (CryptoAPI) m\u1eabu chu\u1ea9n ""abc""")
    Check IsSha256(Sha256Hex(b)) And Not IsSha256(String(63, "a")) And Not IsSha256(String(63, "a") & "g"), U("ki\u1ec3m tra d\u1ea1ng SHA-256")
End Sub

Private Sub TestPrint()
    Dim days(0 To 3) As NKTC_SourceDay, picked() As Long, n As Long, msg As String, i As Long
    Const D As Long = 46188
    days(0).DayNum = D: days(0).ValueRow = 2: days(0).RowNumber = 2
    days(1).DayNum = D + 1: days(1).ValueRow = 3: days(1).RowNumber = 3
    days(2).DayNum = D + 3: days(2).ValueRow = 4: days(2).RowNumber = 4
    days(3).DayNum = D + 7: days(3).ValueRow = 5: days(3).RowNumber = 5

    msg = SelectPrintDays(days, 4, PRINT_ALL, 0, 0, 0, picked, n)
    Check msg = "" And n = 4, U("in t\u1ea5t c\u1ea3: 4 ng\u00e0y")
    Check picked(0) = 0 And picked(3) = 3, U("in t\u1ea5t c\u1ea3: ch\u1ec9 s\u1ed1 t\u0103ng d\u1ea7n")

    msg = SelectPrintDays(days, 4, PRINT_ONE, D + 3, 0, 0, picked, n)
    Check msg = "" And n = 1, U("m\u1ed9t ng\u00e0y c\u00f3 trong danh s\u00e1ch")
    Eq picked(0), 2, U("m\u1ed9t ng\u00e0y: ch\u1ec9 s\u1ed1")
    msg = SelectPrintDays(days, 4, PRINT_ONE, D + 2, 0, 0, picked, n)
    Eq msg, U("Kh\u00f4ng c\u00f3 ng\u00e0y 17/06/2026 trong v\u00f9ng ngu\u1ed3n."), U("m\u1ed9t ng\u00e0y kh\u00f4ng c\u00f3")
    Eq n, 0, U("m\u1ed9t ng\u00e0y kh\u00f4ng c\u00f3: s\u1ed1 ng\u00e0y ch\u1ecdn")
    Eq SelectPrintDays(days, 4, PRINT_ONE, 0, 0, 0, picked, n), U("Nh\u1eadp ng\u00e0y c\u1ea7n in (dd/mm/yyyy)."), U("m\u1ed9t ng\u00e0y: ch\u01b0a nh\u1eadp")

    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 1, D + 3, 0, picked, n)
    Check msg = "" And n = 2, U("kho\u1ea3ng ng\u00e0y th\u01b0\u1eddng")
    Check picked(0) = 1 And picked(1) = 2, U("kho\u1ea3ng ng\u00e0y: ch\u1ec9 s\u1ed1")
    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 7, D, 0, picked, n)
    Check msg = "" And n = 4, U("kho\u1ea3ng \u0111\u1ea3o ng\u01b0\u1ee3c (from > to)")
    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 4, D + 6, 0, picked, n)
    Eq msg, U("Kh\u00f4ng c\u00f3 ng\u00e0y n\u00e0o t\u1eeb 19/06/2026 \u0111\u1ebfn 21/06/2026 trong v\u00f9ng ngu\u1ed3n."), U("kho\u1ea3ng r\u1ed7ng")
    Eq n, 0, U("kho\u1ea3ng r\u1ed7ng: s\u1ed1 ng\u00e0y ch\u1ecdn")
    Eq SelectPrintDays(days, 4, PRINT_RANGE, D, 0, 0, picked, n), U("Nh\u1eadp \u0111\u1ee7 T\u1eeb ng\u00e0y v\u00e0 \u0110\u1ebfn ng\u00e0y (dd/mm/yyyy)."), U("kho\u1ea3ng thi\u1ebfu \u0110\u1ebfn ng\u00e0y")
    Eq SelectPrintDays(days, 4, PRINT_RANGE, 0, D, 0, picked, n), U("Nh\u1eadp \u0111\u1ee7 T\u1eeb ng\u00e0y v\u00e0 \u0110\u1ebfn ng\u00e0y (dd/mm/yyyy)."), U("kho\u1ea3ng thi\u1ebfu T\u1eeb ng\u00e0y")

    msg = SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, D + 7, picked, n)
    Check msg = "" And n = 1, U("ng\u00e0y \u0111ang hi\u1ec3n th\u1ecb c\u00f3")
    Eq picked(0), 3, U("ng\u00e0y \u0111ang hi\u1ec3n th\u1ecb: ch\u1ec9 s\u1ed1")
    Eq SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, 0, picked, n), _
        U("Bi\u1ec3u m\u1eabu ch\u01b0a hi\u1ec3n th\u1ecb ng\u00e0y n\u00e0o. B\u1ea5m m\u1ed9t n\u00fat Chuy\u1ec3n ng\u00e0y tr\u01b0\u1edbc, ho\u1eb7c ch\u1ecdn ng\u00e0y c\u1ea7n in."), U("ch\u01b0a hi\u1ec3n th\u1ecb ng\u00e0y n\u00e0o")
    Eq SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, D + 2, picked, n), _
        U("Bi\u1ec3u m\u1eabu ch\u01b0a hi\u1ec3n th\u1ecb ng\u00e0y n\u00e0o. B\u1ea5m m\u1ed9t n\u00fat Chuy\u1ec3n ng\u00e0y tr\u01b0\u1edbc, ho\u1eb7c ch\u1ecdn ng\u00e0y c\u1ea7n in."), U("ng\u00e0y \u0111ang hi\u1ec3n th\u1ecb kh\u00f4ng c\u00f2n trong danh s\u00e1ch")

    Eq SelectPrintDays(days, 0, PRINT_ALL, 0, 0, 0, picked, n), U("V\u00f9ng ngu\u1ed3n kh\u00f4ng c\u00f3 ng\u00e0y h\u1ee3p l\u1ec7 \u0111\u1ec3 in."), U("kh\u00f4ng c\u00f3 ng\u00e0y n\u00e0o")
    Eq n, 0, U("kh\u00f4ng c\u00f3 ng\u00e0y n\u00e0o: s\u1ed1 ng\u00e0y ch\u1ecdn")
    Eq SelectPrintDays(days, 4, 9, 0, 0, 0, picked, n), U("Ch\u1ebf \u0111\u1ed9 in kh\u00f4ng h\u1ee3p l\u1ec7."), U("ch\u1ebf \u0111\u1ed9 sai (9)")
    Eq SelectPrintDays(days, 4, 0, 0, 0, 0, picked, n), U("Ch\u1ebf \u0111\u1ed9 in kh\u00f4ng h\u1ee3p l\u1ec7."), U("ch\u1ebf \u0111\u1ed9 sai (0)")

    Eq PrintPdfFileName(D, D), "NKTC_15-06-2026.pdf", U("t\u00ean PDF m\u1ed9t ng\u00e0y")
    Eq PrintPdfFileName(D, D + 199), "NKTC_15-06-2026_den_31-12-2026.pdf", U("t\u00ean PDF nhi\u1ec1u ng\u00e0y")
    Eq PrintSheetName(D), "Ngay_15-06-2026", U("t\u00ean sheet t\u1ea1m")
    Check Len(PrintSheetName(D)) <= 31, U("t\u00ean sheet t\u1ea1m <= 31 k\u00fd t\u1ef1")
    For i = 1 To Len(PrintSheetName(D))
        Check InStr("\/?*[]:", Mid$(PrintSheetName(D), i, 1)) = 0, U("t\u00ean sheet t\u1ea1m kh\u00f4ng c\u00f3 k\u00fd t\u1ef1 c\u1ea5m")
    Next i
End Sub
