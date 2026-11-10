Attribute VB_Name = "modNKTC_Tests"
'==============================================================================
' NHẬT KÝ TC – kiểm thử phần lõi (tương ứng các file *.test.ts của bản Office.js).
' Chạy: Application.Run "NhatKyThiCong_Core.xlam!modNKTC_Tests.NKTC_RunTests" -> "PASS n" hoặc "FAIL ...".
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
    If Not ok Then name = name & " (được: " & CStr(actual) & " | cần: " & CStr(expected) & ")"
    Check ok, name
End Sub

' Mảng 2 chiều 1-based từ các dòng (mỗi dòng là Array(...)).
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

' Chạy một nhóm kiểm thử; lỗi khi chạy được ghi vào kết quả (không hiện hộp thoại) rồi chạy tiếp nhóm sau.
Private Sub RunOne(ByVal idx As Long, ByVal testName As String)
    On Error GoTo Loi
    mLast = "(chưa có phép kiểm tra nào)"
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
    mLog = mLog & vbLf & "  x " & testName & ": lỗi " & Err.Number & " – " & Err.Description & _
        " (sau phép kiểm tra: " & mLast & ")"
End Sub

Private Sub TestText()
    Eq U("Ngày"), "Ngày", "U giải mã"
    Eq RemoveDiacritics("Ngày Bắt Đầu"), "ngay bat dau", "bỏ dấu 1"
    Eq RemoveDiacritics("NỘI DUNG CÔNG VIỆC"), "noi dung cong viec", "bỏ dấu chữ hoa"
    Eq RemoveDiacritics("Thời tiết, Ước lượng, Kỹ thuật"), "thoi tiet, uoc luong, ky thuat", "bỏ dấu 3"
    Eq NormalizeHeader("  Ngày" & vbLf & "bắt   đầu (*) "), "ngay bat dau", "chuẩn hóa tiêu đề"
    Eq NormalizeHeader("Ban chỉ huy - CBKT"), "ban chi huy cbkt", "chuẩn hóa tiêu đề 2"
    Eq NormalizeHeaderName("  THỜI   tiết "), NormalizeHeaderName("Thời tiết"), "tên tiêu đề không phân biệt hoa/khoảng trắng"
    Check NormalizeHeaderName("Thoi tiet") <> NormalizeHeaderName("Thời tiết"), "tên tiêu đề giữ dấu"
    Eq CellText(12), "12", "CellText số"
    Check IsBlankValue(Empty) And IsBlankValue("  ") And Not IsBlankValue(0), "IsBlankValue"
End Sub

Private Sub TestDates()
    Const NOV1 As Long = 46327
    Eq DayFromCalendar(2026, 11, 1), NOV1, "DATE(2026,11,1)"
    Eq DayFromCalendar(1900, 3, 1), MIN_DAY, "MIN_DAY"
    Eq DayFromCalendar(9999, 12, 31), MAX_DAY, "MAX_DAY"
    Eq DayFromCalendar(2026, 2, 29), 0, "29/02/2026 không tồn tại"
    Eq DayFromCalendar(2026, 4, 31), 0, "31/04 không tồn tại"
    Eq DayFromCalendar(1900, 2, 28), 0, "trước MIN_DAY"
    Check DayFromCalendar(2028, 2, 29) > 0, "29/02/2028"
    Eq FormatDay(NOV1), "01/11/2026", "FormatDay"
    Eq FormatDay(NOV1 + 60), "31/12/2026", "FormatDay +60"
    Dim ok As Variant
    For Each ok In Array("01/11/2026", "1/11/2026", "01-11-2026", "01.11.2026", "  01/11/2026  ", "2026-11-01", "01/11-2026")
        Eq ParseDateText(ok), NOV1, "đọc " & ok
    Next ok
    For Each ok In Array("31/02/2026", "11/2026", "01/11/26", "abc", "", "01/11/2026 08:00")
        Eq ParseDateText(ok), 0, "từ chối " & ok
    Next ok
    Eq ToDayNumber(CDbl(NOV1) + 0.75), NOV1, "bỏ phần giờ"
    Eq ToDayNumber("01/11/2026"), NOV1, "ToDayNumber chữ"
    Eq ToDayNumber(0), 0, "0 không phải ngày"
    Eq ToDayNumber(-5), 0, "số âm"
    Eq ToDayNumber(CDbl(MAX_DAY) + 1), 0, "quá MAX_DAY"
    Eq ToDayNumber(True), 0, "logic"
    Eq ToDayNumber(Empty), 0, "trống"
    Eq ToDayNumber(NOV1 - 1462, True), NOV1, "hệ 1904"
    Eq ToWorkbookSerial(NOV1, True), NOV1 - 1462, "ghi hệ 1904"
End Sub

Private Sub TestDatePattern()
    Const D As Long = 46188 ' 15/06/2026, thứ Hai
    Eq FillDatePattern("Ngày ..... tháng ..... năm ..........", D), "Ngày 15 tháng 06 năm 2026", "rải ngày"
    Eq FillDatePattern("Ngày 01 tháng 01 năm 2025", D + 1), "Ngày 16 tháng 06 năm 2026", "điền lại"
    Eq FillDatePattern("Hà Nội, ngày … tháng … năm …", D), "Hà Nội, ngày 15 tháng 06 năm 2026", "giữ phần trước"
    Eq FillDatePattern("NGÀY __ THÁNG __ NĂM ____", D), "NGÀY 15 THÁNG 06 NĂM 2026", "giữ chữ hoa"
    Eq FillDatePattern("Nội dung công việc", D), "", "không phải mẫu"
    Eq FillDatePattern("năm tháng ngày", D), "", "sai thứ tự"
    Check IsDatePattern("Ngày ... tháng ... năm ...") And Not IsDatePattern("Thời tiết"), "IsDatePattern"
    Eq TenThu(D), "Thứ Hai", "thứ Hai"
    Eq TenThu(D + 6), "Chủ nhật", "Chủ nhật"
End Sub

Private Sub TestColumns()
    Dim c As Long, s As Long, e As Long
    DetectColumns Array("STT", "Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"), c, s, e
    Check c = 1 And s = 2 And e = 3, "tiêu đề chuẩn"
    DetectColumns Array("TÊN CÔNG VIỆC", "Tu ngay", "Den ngay"), c, s, e
    Check c = 0 And s = 1 And e = 2, "không dấu, viết hoa"
    DetectColumns Array("Hạng mục công việc", "Thời gian bắt đầu", "Thời gian kết thúc"), c, s, e
    Check c = 0 And s = 1 And e = 2, "chứa cụm từ"
    DetectColumns Array("TÊN HẠNG MỤC/ CÔNG TÁC", "BẮT ĐẦU", "KẾT THÚC"), c, s, e
    Check c = 0 And s = 1 And e = 2, "hạng mục / công tác"
    DetectColumns Array("STT", "Hạng mục", "Từ ngày", "Đến ngày"), c, s, e
    Check c = 1 And s = 2 And e = 3, "từ ngày / đến ngày"
    DetectColumns Array("Nội dung nghiệm thu", "Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"), c, s, e
    Check c = 1 And s = 2 And e = 3, "không nhầm cột nội dung phụ"
    DetectColumns Array(1, Empty, "Ghi chú"), c, s, e
    Check c = -1 And s = -1 And e = -1, "không nhận diện được"
    Check IsCompleteColumns(0, 1, 2) And Not IsCompleteColumns(0, 1, -1) And Not IsCompleteColumns(0, 1, 1), "IsCompleteColumns"
End Sub

Private Sub TestDailyLog()
    Dim items() As NKTC_WorkItem, n As Long, issues() As NKTC_RowIssue, ni As Long, blank As Long
    Dim entries() As NKTC_DailyEntry, ne As Long, f As Long, l As Long, msg As String
    Dim v As Variant
    Const D As Long = 46188
    v = Grid(Array("Nội dung", "Bắt đầu", "Kết thúc"), _
             Array("- Đào đất", CDbl(D), CDbl(D + 2)), _
             Array("Đổ bê tông", "16/06/2026", Empty), _
             Array(Empty, Empty, Empty), _
             Array("", CDbl(D), CDbl(D)), _
             Array("Sai ngày", "abc", Empty), _
             Array("Ngược", CDbl(D + 3), CDbl(D)))
    ParseWorkItems v, 0, 1, 2, 1, False, items, n, issues, ni, blank
    Eq n, 2, "số công việc hợp lệ"
    Eq ni, 3, "số dòng lỗi"
    Eq blank, 1, "dòng trống"
    Eq items(1).EndDay, D + 1, "thiếu ngày kết thúc -> 1 ngày"
    Eq issues(0).RowNumber, 5, "số dòng lỗi đầu"
    Eq issues(0).Reason, "Thiếu nội dung công việc.", "lý do thiếu nội dung"
    Eq issues(1).Reason, "Ngày bắt đầu không hợp lệ.", "lý do ngày sai"
    Eq issues(2).Reason, "Ngày kết thúc trước ngày bắt đầu.", "lý do ngược ngày"
    msg = BuildDailyLog(items, n, entries, ne, f, l)
    Eq msg, "", "BuildDailyLog thành công"
    Eq ne, 3, "số ngày"
    Eq entries(1).Content, "- Đào đất" & vbLf & "Đổ bê tông", "gộp nội dung, giữ nguyên dấu -"
    Eq entries(1).ItemCount, 2, "đếm công việc"
    items(0).EndDay = D + MAX_SPAN_DAYS
    msg = BuildDailyLog(items, n, entries, ne, f, l)
    Check InStr(msg, "vượt giới hạn") > 0, "chặn khoảng quá dài"
    msg = BuildDailyLog(items, 0, entries, ne, f, l)
    Eq msg, "Không có công việc hợp lệ để xuất.", "không có công việc"
End Sub

Private Sub TestSheetName()
    Dim names As New Collection
    Eq UniqueSheetName(names), "NK_TheoNgay", "tên gốc"
    names.Add "nk_theongay": names.Add "NK_TheoNgay_2"
    Eq UniqueSheetName(names), "NK_TheoNgay_3", "tên tiếp theo, không phân biệt hoa/thường"
End Sub

Private Sub TestDays()
    Dim v As Variant, days() As NKTC_SourceDay, n As Long, issues As Collection
    Const D As Long = 46188
    v = Grid(Array("Ngày", "Thời tiết"), Array(CDbl(D + 1), "a"), Array(CDbl(D), "b"), Array(Empty, "c"), _
             Array("xx", "d"), Array(CDbl(D), "e"))
    Eq DetectDateColumn(v), 0, "cột Ngày theo tiêu đề"
    BuildDayIndex v, 0, 1, False, days, n, issues
    Eq n, 2, "số ngày"
    Eq days(0).DayNum, D, "sắp xếp tăng dần"
    Eq days(0).RowNumber, 3, "số dòng"
    Eq issues.Count, 2, "báo ngày sai và trùng"
    Check InStr(issues(2), "Trùng ngày 15/06/2026 với dòng 3") > 0, "nội dung báo trùng"
    Eq IndexOfDay(days, n, D + 1), 1, "IndexOfDay"
    v = Grid(Array("A", "Ghi chú ngày"), Array("x", CDbl(D)))
    Eq DetectDateColumn(v), 1, "tiêu đề có từ ngay"
    v = Grid(Array("A", "B"), Array("x", CDbl(D)), Array("y", "16/06/2026"))
    Eq DetectDateColumn(v), 1, "cột toàn ngày"
End Sub

Private Sub TestNavigation()
    Dim v As Variant, days() As NKTC_SourceDay, n As Long, issues As Collection, r As NKTC_DayResolution
    Const D As Long = 46188
    v = Grid(Array("Ngày"), Array(CDbl(D)), Array(CDbl(D + 1)), Array(CDbl(D + 2)))
    BuildDayIndex v, 0, 1, False, days, n, issues
    r = ResolveDay(days, n, 0, CMD_FIRST): Check r.Ok And r.TargetIndex = 0, "ngày đầu"
    r = ResolveDay(days, n, 0, CMD_LAST): Check r.Ok And r.TargetIndex = 2, "ngày cuối"
    r = ResolveDay(days, n, D, CMD_NEXT): Check r.Ok And r.TargetIndex = 1, "ngày sau"
    r = ResolveDay(days, n, D + 2, CMD_NEXT): Check (Not r.Ok) And r.IsInfo And InStr(r.Message, "Đang ở ngày cuối cùng (17/06/2026)") > 0, "đã ở cuối"
    r = ResolveDay(days, n, D, CMD_PREV): Check (Not r.Ok) And r.IsInfo And InStr(r.Message, "Đang ở ngày đầu tiên") > 0, "đã ở đầu"
    r = ResolveDay(days, n, 0, CMD_NEXT): Check r.Ok And r.TargetIndex = 0 And r.Notice = "", "chưa hiển thị -> ngày đầu"
    r = ResolveDay(days, n, D + 9, CMD_PREV): Check r.Ok And r.TargetIndex = 0 And InStr(r.Notice, "Không còn ngày") > 0, "ngày cũ không còn"
    r = ResolveDay(days, 0, 0, CMD_FIRST): Check (Not r.Ok) And (Not r.IsInfo), "không có ngày"
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
    oldH = Strs("Ngày", "Nội dung", "Thời tiết", "Ghi chú", "Ghi chú")
    newH = Strs("Ngày", "Nội dung", "Số công nhân", "Thời tiết", "Ghi chú", "Ghi chú")
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = 1 And m(2) = 3 And m(3) = 4 And m(4) = 5, "chèn cột giữa"
    oldH = Strs("Ghi chú", "A", "Ghi chú")
    newH = Strs("Ghi chú", "Ghi chú", "A")
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = 2 And m(2) = 1, "trùng tên theo thứ tự"
    oldH = Strs("Ngày", "", "Thời tiết")
    newH = Strs("ngày ", "", "Thời tiết (sáng)")
    m = MatchHeaders(oldH, newH)
    Check m(0) = 0 And m(1) = -1 And m(2) = -1, "không phân biệt hoa/khoảng trắng; trống và đổi tên không khớp"
End Sub

Private Sub TestWriteValue()
    Const D As Long = 46188
    Eq ToCellInput("=SUM(A1)"), "'=SUM(A1)", "chuỗi thêm dấu '"
    Eq ToCellInput(""), "", "chuỗi rỗng"
    Eq ToCellInput(5), 5, "số giữ nguyên"
    Eq ToCellInput(Empty), "", "trống"
    Eq WriteValue(CDbl(D), True, "Ngày ..... tháng ..... năm ..........", False), "'Ngày 15 tháng 06 năm 2026", "cột Ngày vào ô mẫu ngày"
    Eq WriteValue(CDbl(D), True, "", False), CDbl(D), "cột Ngày vào ô trống -> số"
    Eq WriteValue("Nắng", False, "Ngày ... tháng ... năm ...", False), "'Nắng", "cột khác giữ nguyên"
End Sub

Private Sub TestValidate()
    Dim maps(0 To 2) As NKTC_Mapping, errs As Collection
    maps(0).Column = 0: maps(0).SheetName = "Mẫu NKTC": maps(0).Address = "E2"
    maps(1).Column = 1: maps(1).SheetName = "mẫu nktc": maps(1).Address = "e2"
    maps(2).Column = 9: maps(2).SheetName = "Mẫu NKTC": maps(2).Address = "C6"
    Set errs = ValidateMappings(maps, 3, 5)
    Eq errs.Count, 2, "ô trùng + cột ngoài vùng"
    Eq MappingKey(maps, 2), "mẫu nktc!E2|mẫu nktc!E2", "MappingKey"
    Set errs = ValidateMappings(maps, 0, 5)
    Eq errs.Count, 1, "chưa gán"
End Sub

Private Sub TestDialogHelpers()
    Dim cc As Long, cs As Long, ce As Long
    Eq ColumnLetter(0), "A", "cột A"
    Eq ColumnLetter(25), "Z", "cột Z"
    Eq ColumnLetter(26), "AA", "cột AA"
    Eq ColumnLetter(16383), "XFD", "cột cuối XFD"
    Eq ColumnOptionLabel("  Nội dung" & vbLf & "công việc ", 2), "Cột C – Nội dung công việc", "nhãn cột"
    Eq ColumnOptionLabel(Empty, 0), "Cột A – (không có tiêu đề)", "nhãn cột trống"
    Eq ColumnOptionLabel(2026, 1), "Cột B – 2026", "nhãn cột số"
    Eq Len(ShortLabel(String(80, "x"))), 60, "nhãn dài cắt còn 60"
    Eq ParseRowCount(" 10 "), 10, "số dòng 10"
    Eq ParseRowCount("1000"), 1000, "số dòng tối đa"
    Eq ParseRowCount("1001"), 0, "quá số dòng tối đa"
    Eq ParseRowCount("0"), 0, "số dòng 0"
    Eq ParseRowCount("5a"), 0, "số dòng có chữ"
    Eq ParseRowCount(""), 0, "số dòng trống"
    ' Bảng do Tạo bảng sinh ra phải được Xuất theo ngày nhận diện đủ 3 cột (R1.12).
    DetectColumns InputTableHeaders(), cc, cs, ce
    Check cc = 1 And cs = 2 And ce = 3, "Tạo bảng -> nhận diện đủ cột"
End Sub

Private Sub TestReference()
    Dim sh As String, ad As String
    Eq NormalizeCellAddress(" $e$2 "), "E2", "địa chỉ ô bỏ $, viết hoa"
    Eq NormalizeCellAddress("XFD1048576"), "XFD1048576", "ô cuối sheet"
    Eq NormalizeCellAddress("XFE1"), "", "cột vượt XFD"
    Eq NormalizeCellAddress("A0"), "", "dòng 0"
    Eq NormalizeCellAddress("A1:B2"), "", "vùng không phải một ô"
    Check ParseCellReference("'Mẫu NKTC'!$E$2", sh, ad) And sh = "Mẫu NKTC" And ad = "E2", "ô có tên sheet trong nháy"
    Check ParseCellReference("Mẫu NKTC!e2", sh, ad) And sh = "Mẫu NKTC" And ad = "E2", "ô có tên sheet không nháy"
    Check ParseCellReference("E2", sh, ad) And sh = "" And ad = "E2", "ô không ghi sheet"
    Check ParseCellReference("'It''s'!B3", sh, ad) And sh = "It's", "tên sheet có dấu nháy"
    Check Not ParseCellReference("A1:B2", sh, ad), "vùng nhiều ô bị từ chối"
    Check Not ParseCellReference("Sh[1]!A1", sh, ad), "tên sheet có ký tự cấm"
    Check Not ParseCellReference("", sh, ad), "ô trống"
    Eq FormatReference("Mẫu NKTC", "E2"), "Mẫu NKTC!E2", "hiển thị ô"
    Eq QuotedReference("It's", "B3"), "'It''s'!B3", "tham chiếu cho Excel"
End Sub

Private Sub TestFormRows()
    Dim cols() As Long, targets() As String, maps() As NKTC_Mapping, n As Long, errors As Collection, h() As String
    ReDim cols(0 To 3): ReDim targets(0 To 3)
    cols(0) = 0: targets(0) = "Mẫu NKTC!E2"
    cols(1) = -1: targets(1) = ""
    cols(2) = 1: targets(2) = "e2"
    cols(3) = 2: targets(3) = "x"
    Set errors = BuildMappingsFromRows(cols, targets, 4, 3, "Mẫu NKTC", maps, n)
    Eq errors.Count, 2, "số lỗi dòng gán"
    Eq errors(1), "Dòng 3: ô Mẫu NKTC!E2 đã được gán ở dòng 1.", "trùng ô (ô không ghi sheet thuộc sheet mặc định)"
    Eq errors(2), "Dòng 4: ""x"" không phải một ô hợp lệ (ví dụ Mẫu NKTC!E2).", "ô không hợp lệ"
    Eq n, 0, "có lỗi thì không có gán"
    cols(2) = -1: targets(2) = "F3"
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, "Mẫu NKTC", maps, n)
    Eq errors(1), "Dòng 3: chưa chọn cột nguồn.", "thiếu cột"
    cols(2) = 1: targets(2) = ""
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, "Mẫu NKTC", maps, n)
    Eq errors(1), "Dòng 3: chưa chọn ô trên form.", "thiếu ô"
    targets(2) = "F3"
    Set errors = BuildMappingsFromRows(cols, targets, 3, 3, "Form", maps, n)
    Eq errors.Count, 0, "dòng gán hợp lệ"
    Check n = 2 And maps(1).Column = 1 And maps(1).SheetName = "Form" And maps(1).Address = "F3", "gán dùng sheet mặc định"
    Set errors = BuildMappingsFromRows(cols, targets, 0, 3, "Form", maps, n)
    Check errors.Count = 1 And InStr(errors(1), "Chưa có dòng gán") > 0, "chưa có dòng gán"
    h = Strs("a", " ", "b", "")
    Eq BlankHeaderList(h, 4), "2, 4", "tiêu đề trống"
    h = Strs("a", "b")
    Eq BlankHeaderList(h, 2), "", "không có tiêu đề trống"
End Sub

Private Sub TestConfig()
    Dim cfg As NKTC_Config, back As NKTC_Config, maps() As NKTC_Mapping, s As String, msg As String, h() As String
    s = "a\b" & vbTab & "c" & vbLf & "d" & vbCr & "e"
    Eq UnescapeField(EscapeField(s)), s, "mã hóa trường giữ nguyên nội dung"
    Check InStr(EscapeField(s), vbTab) = 0 And InStr(EscapeField(s), vbLf) = 0, "trường mã hóa không chứa Tab/xuống dòng"
    Eq XmlText("a<b&c>" & vbTab & vbLf), "a&lt;b&amp;c&gt;&#9;&#10;", "chữ trong XML"

    h = Strs("Ngày", "Thời tiết" & vbLf & "(sáng)", "Nội dung")
    SetHeaders cfg, h, 3
    ReDim maps(0 To 1)
    maps(0).Column = 1: maps(0).SheetName = "Mẫu NKTC": maps(0).Address = "E5"
    maps(1).Column = 2: maps(1).SheetName = "Mẫu NKTC": maps(1).Address = "B9"
    SetMaps cfg, maps, 2
    cfg.SourceSheet = "NK_TheoNgay": cfg.SourceAddress = "A1:C40"
    cfg.DateColumn = 0: cfg.CurrentDay = 46188: cfg.ConfirmedKey = MappingKey(maps, 2)
    Check ParseConfig(SerializeConfig(cfg), back), "đọc lại cấu hình đã lưu"
    Check back.SourceSheet = "NK_TheoNgay" And back.SourceAddress = "A1:C40" And back.DateColumn = 0 And _
          back.CurrentDay = 46188 And back.ConfirmedKey = cfg.ConfirmedKey, "cấu hình: vùng nguồn, ngày, xác nhận"
    Check back.HeaderCount = 3 And back.Headers(1) = "Thời tiết" & vbLf & "(sáng)", "cấu hình: tiêu đề nhiều dòng"
    Check back.MapCount = 2 And back.Maps(1).Column = 2 And back.Maps(1).SheetName = "Mẫu NKTC" And _
          back.Maps(1).Address = "B9", "cấu hình: dòng gán"
    Check Not ParseConfig("abc", back), "cấu hình hỏng bị bỏ qua"
    Check Not ParseConfig(Replace(SerializeConfig(cfg), "date" & vbTab & "0", "date" & vbTab & "7"), back), _
        "cột Ngày ngoài vùng bị từ chối"

    ' Khớp lại theo tên tiêu đề: chèn cột giữa -> cột đã gán dịch theo.
    h = Strs("Ngày", "Thời tiết", "Nội dung")
    SetHeaders cfg, h, 3
    h = Strs("Ngày", "Số công nhân", "THỜI  TIẾT", "Nội dung")
    msg = RemapConfig(cfg, h, 4, "NK!A1:D9")
    Eq msg, "", "khớp lại sau khi chèn cột"
    Check cfg.DateColumn = 0 And cfg.Maps(0).Column = 2 And cfg.Maps(1).Column = 3 And cfg.HeaderCount = 4, _
        "vị trí cột sau khi chèn cột"
    h = Strs("Ngày", "Nội dung")
    msg = RemapConfig(cfg, h, 2, "NK!A1:B9")
    Check InStr(msg, "Không tìm thấy cột ""THỜI  TIẾT"" (gán vào Mẫu NKTC!E5)") > 0, "báo cột bị xóa kèm ô đích"
    Check cfg.Maps(0).Column = 2, "cột thiếu thì không đổi cấu hình"
End Sub

Private Sub TestUpdate()
    Dim ver As String, f As String, sha As String, size As Long, notes As String, b() As Byte
    ParseManifest "version=3.1.0" & vbCrLf & "file=NhatKyThiCong_Core.xlam" & vbCrLf & "sha256=" & String(64, "A") & vbLf & _
        "size=12345" & vbLf & "notes=Sửa lỗi" & vbLf, ver, f, sha, size, notes
    Check ver = "3.1.0" And f = "NhatKyThiCong_Core.xlam" And sha = String(64, "a") And size = 12345 And notes = "Sửa lỗi", _
        "đọc version.txt"
    ParseManifest "version=3.1" & vbLf & "file=..\x.xlam" & vbLf & "sha256=xyz", ver, f, sha, size, notes
    Check ver = "3.1" And f = "" And sha = "", "version.txt: tên tệp có đường dẫn, SHA-256 sai bị bỏ"
    ParseManifest "file=https://x/y.xlam" & vbLf & "version=3.0.0-beta", ver, f, sha, size, notes
    Check ver = "" And f = "", "version.txt: phiên bản sai, tên tệp là URL bị bỏ"
    Eq CompareVersions("3.10.0", "3.9.0"), 1, "so phiên bản 3.10 > 3.9"
    Eq CompareVersions("3.0.0", "3.0.0"), 0, "so phiên bản bằng"
    Eq CompareVersions("3.0", "3.0.1"), -1, "so phiên bản thiếu phần"
    Check IsVersion("3.0.0") And Not IsVersion("3") And Not IsVersion("3.a.0") And Not IsVersion("3..0"), "kiểm tra dạng phiên bản"
    Check IsAllowedUpdateUrl("https://khoacrab-97.github.io/x/") And IsAllowedUpdateUrl("http://127.0.0.1:8080/"), "địa chỉ cập nhật hợp lệ"
    Check Not IsAllowedUpdateUrl("http://example.com/") And Not IsAllowedUpdateUrl("https://u:p@example.com/") And _
          Not IsAllowedUpdateUrl("http://localhost.evil.com/"), "địa chỉ cập nhật bị từ chối"
    b = StrConv("abc", vbFromUnicode)
    Eq Sha256Hex(b), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256 (CryptoAPI) mẫu chuẩn ""abc"""
    Check IsSha256(Sha256Hex(b)) And Not IsSha256(String(63, "a")) And Not IsSha256(String(63, "a") & "g"), "kiểm tra dạng SHA-256"
End Sub

Private Sub TestPrint()
    Dim days(0 To 3) As NKTC_SourceDay, picked() As Long, n As Long, msg As String, i As Long
    Const D As Long = 46188
    days(0).DayNum = D: days(0).ValueRow = 2: days(0).RowNumber = 2
    days(1).DayNum = D + 1: days(1).ValueRow = 3: days(1).RowNumber = 3
    days(2).DayNum = D + 3: days(2).ValueRow = 4: days(2).RowNumber = 4
    days(3).DayNum = D + 7: days(3).ValueRow = 5: days(3).RowNumber = 5

    msg = SelectPrintDays(days, 4, PRINT_ALL, 0, 0, 0, picked, n)
    Check msg = "" And n = 4, "in tất cả: 4 ngày"
    Check picked(0) = 0 And picked(3) = 3, "in tất cả: chỉ số tăng dần"

    msg = SelectPrintDays(days, 4, PRINT_ONE, D + 3, 0, 0, picked, n)
    Check msg = "" And n = 1, "một ngày có trong danh sách"
    Eq picked(0), 2, "một ngày: chỉ số"
    msg = SelectPrintDays(days, 4, PRINT_ONE, D + 2, 0, 0, picked, n)
    Eq msg, "Không có ngày 17/06/2026 trong vùng nguồn.", "một ngày không có"
    Eq n, 0, "một ngày không có: số ngày chọn"
    Eq SelectPrintDays(days, 4, PRINT_ONE, 0, 0, 0, picked, n), "Nhập ngày cần in (dd/mm/yyyy).", "một ngày: chưa nhập"

    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 1, D + 3, 0, picked, n)
    Check msg = "" And n = 2, "khoảng ngày thường"
    Check picked(0) = 1 And picked(1) = 2, "khoảng ngày: chỉ số"
    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 7, D, 0, picked, n)
    Check msg = "" And n = 4, "khoảng đảo ngược (from > to)"
    msg = SelectPrintDays(days, 4, PRINT_RANGE, D + 4, D + 6, 0, picked, n)
    Eq msg, "Không có ngày nào từ 19/06/2026 đến 21/06/2026 trong vùng nguồn.", "khoảng rỗng"
    Eq n, 0, "khoảng rỗng: số ngày chọn"
    Eq SelectPrintDays(days, 4, PRINT_RANGE, D, 0, 0, picked, n), "Nhập đủ Từ ngày và Đến ngày (dd/mm/yyyy).", "khoảng thiếu Đến ngày"
    Eq SelectPrintDays(days, 4, PRINT_RANGE, 0, D, 0, picked, n), "Nhập đủ Từ ngày và Đến ngày (dd/mm/yyyy).", "khoảng thiếu Từ ngày"

    msg = SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, D + 7, picked, n)
    Check msg = "" And n = 1, "ngày đang hiển thị có"
    Eq picked(0), 3, "ngày đang hiển thị: chỉ số"
    Eq SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, 0, picked, n), _
        "Biểu mẫu chưa hiển thị ngày nào. Bấm một nút Chuyển ngày trước, hoặc chọn ngày cần in.", "chưa hiển thị ngày nào"
    Eq SelectPrintDays(days, 4, PRINT_CURRENT, 0, 0, D + 2, picked, n), _
        "Biểu mẫu chưa hiển thị ngày nào. Bấm một nút Chuyển ngày trước, hoặc chọn ngày cần in.", "ngày đang hiển thị không còn trong danh sách"

    Eq SelectPrintDays(days, 0, PRINT_ALL, 0, 0, 0, picked, n), "Vùng nguồn không có ngày hợp lệ để in.", "không có ngày nào"
    Eq n, 0, "không có ngày nào: số ngày chọn"
    Eq SelectPrintDays(days, 4, 9, 0, 0, 0, picked, n), "Chế độ in không hợp lệ.", "chế độ sai (9)"
    Eq SelectPrintDays(days, 4, 0, 0, 0, 0, picked, n), "Chế độ in không hợp lệ.", "chế độ sai (0)"

    Eq PrintPdfFileName(D, D), "NKTC_15-06-2026.pdf", "tên PDF một ngày"
    Eq PrintPdfFileName(D, D + 199), "NKTC_15-06-2026_den_31-12-2026.pdf", "tên PDF nhiều ngày"
    Eq PrintSheetName(D), "Ngay_15-06-2026", "tên sheet tạm"
    Check Len(PrintSheetName(D)) <= 31, "tên sheet tạm <= 31 ký tự"
    For i = 1 To Len(PrintSheetName(D))
        Check InStr("\/?*[]:", Mid$(PrintSheetName(D), i, 1)) = 0, "tên sheet tạm không có ký tự cấm"
    Next i
End Sub
