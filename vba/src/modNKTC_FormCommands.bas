Attribute VB_Name = "modNKTC_FormCommands"
'==============================================================================
' NHẬT KÝ TC – F2: Chèn form mẫu, Xuất vào form (gán cột), Chuyển ngày, Hướng dẫn.
' Tương ứng src/commands/{formCommands,templateCommands}.ts và src/excel/{settingsStore,templateWriter}.ts.
' Chỉ ghi vào các ô đích đã gán (CLAUDE.md §4); ô đang có công thức/dữ liệu -> hỏi trước khi ghi đè (R2.5).
'==============================================================================
Option Explicit

Private Const CONFIG_NS As String = "urn:nhatkytc:config:1"
Private Const MAX_RISKY_SHOWN As Long = 15

' Tên sheet biểu mẫu (có sẵn trong add-in, lấy từ templates/Mau_NKTC.xlsx khi build).
Public Function FormSheetName() As String
    FormSheetName = "Mẫu NKTC"
End Function

Private Function NotAssignedMessage() As String
    NotAssignedMessage = "Chưa gán cột vào form. Bấm ""Xuất vào form"" (nhóm Form mẫu) để chọn vùng nguồn và gán cột " & _
        "trước — chuyển ngày chỉ hoạt động sau khi đã gán."
End Function

Private Function NeedWorkbook() As Boolean
    If ActiveWorkbook Is Nothing Then
        ShowInfo "Hãy mở một file Excel trước.", vbExclamation
    Else
        NeedWorkbook = True
    End If
End Function

'------------------------------------------------------------------------------
' Cấu hình lưu trong file (CustomXMLPart) — giữ lại khi người dùng lưu file Excel (R2.8).
'------------------------------------------------------------------------------
Public Function LoadConfig(ByVal wb As Workbook, ByRef cfg As NKTC_Config) As Boolean
    Dim parts As Object
    Set parts = wb.CustomXMLParts.SelectByNamespace(CONFIG_NS)
    If parts.Count = 0 Then Exit Function
    LoadConfig = ParseConfig(parts(1).DocumentElement.Text, cfg)
End Function

Public Sub SaveConfig(ByVal wb As Workbook, ByRef cfg As NKTC_Config)
    Dim parts As Object, i As Long
    Set parts = wb.CustomXMLParts.SelectByNamespace(CONFIG_NS)
    For i = parts.Count To 1 Step -1
        parts(i).Delete
    Next i
    wb.CustomXMLParts.Add "<nktc xmlns=""" & CONFIG_NS & """>" & XmlText(SerializeConfig(cfg)) & "</nktc>"
End Sub

' Vùng nguồn của cấu hình, đã tự mở rộng tới hết dữ liệu của sheet. msg: lý do nếu Nothing.
Private Function ConfigSource(ByVal wb As Workbook, ByRef cfg As NKTC_Config, ByRef msg As String) As Range
    Dim ws As Worksheet, r As Range
    Set ws = FindSheet(wb, cfg.SourceSheet)
    If ws Is Nothing Then
        msg = "Không có sheet """ & cfg.SourceSheet & """ (vùng nguồn đã gán) trong file. Bấm ""Xuất vào form"" để chọn lại vùng nguồn."
        Exit Function
    End If
    On Error Resume Next
    Set r = ws.Range(cfg.SourceAddress)
    On Error GoTo 0
    If r Is Nothing Then msg = "Vùng nguồn đã lưu không hợp lệ. Bấm ""Xuất vào form"" để chọn lại.": Exit Function
    Set ConfigSource = ExpandSource(r)
End Function

'------------------------------------------------------------------------------
' Chèn form mẫu: chép sheet "Mẫu NKTC" từ add-in vào cuối file đang mở. Không ghi đè sheet có sẵn.
'------------------------------------------------------------------------------
Public Sub CmdChenForm()
    Dim wb As Workbook, ws As Worksheet
    If Not NeedWorkbook() Then Exit Sub
    Set wb = ActiveWorkbook
    If Not FindSheet(wb, FormSheetName()) Is Nothing Then
        ShowInfo "File đã có sheet """ & FormSheetName() & """. Add-in không ghi đè — đổi tên hoặc xóa sheet đó nếu muốn chèn mẫu mới.", vbExclamation
        Exit Sub
    End If
    If wb.ProtectStructure Then
        ShowInfo "File đang khóa cấu trúc (Protect Workbook) nên không thêm được sheet.", vbExclamation
        Exit Sub
    End If
    ThisWorkbook.Worksheets(FormSheetName()).Copy After:=wb.Sheets(wb.Sheets.Count)
    Set ws = FindSheet(wb, FormSheetName())
    If Not ws Is Nothing Then ws.Activate
    ShowInfo "Đã chèn sheet """ & FormSheetName() & """." & vbLf & vbLf & _
        "Tiếp theo: bấm ""Xuất vào form"" để chọn vùng nhật ký theo ngày và gán cột vào các ô của biểu mẫu."
End Sub

'------------------------------------------------------------------------------
' Chuyển ngày (R2.3–R2.7, R2.15, R2.17).
'------------------------------------------------------------------------------
Public Sub CmdChuyenNgay(ByVal command As Long)
    Dim message As String, style As VbMsgBoxStyle
    If Not NeedWorkbook() Then Exit Sub
    NavigateDay ActiveWorkbook, command, message, style
    If Len(message) > 0 Then ShowInfo message, style
End Sub

' Ghi dữ liệu một ngày lên form: đọc lại nguồn (mới nhất), khớp cột theo tên tiêu đề, chọn ngày, hỏi trước khi
' ghi đè ô có công thức/dữ liệu (một lần cho mỗi bộ ô đích), rồi ghi. True nếu đã ghi; message: thông báo kèm theo.
Public Function NavigateDay(ByVal wb As Workbook, ByVal command As Long, ByRef message As String, _
                            ByRef style As VbMsgBoxStyle, Optional ByVal targetDay As Long = 0) As Boolean
    Dim cfg As NKTC_Config, source As Range, values As Variant, days() As NKTC_SourceDay, dayCount As Long
    Dim res As NKTC_DayResolution
    Dim targets() As Range, texts() As String, i As Long, key As String, risky As String, riskyCount As Long
    Dim raw As Variant, isPattern As Boolean, maps() As NKTC_Mapping, valueRow As Long, c0 As Long, ws As Worksheet

    If Not PrepareSourceDays(wb, cfg, source, values, days, dayCount, message, style) Then Exit Function
    GetMaps cfg, maps
    res = ResolveDay(days, dayCount, cfg.CurrentDay, command, targetDay)
    If Not res.Ok Then
        message = res.Message
        If res.IsInfo Then style = vbInformation
        Exit Function
    End If

    ' Ô đích (ô gộp -> ô đầu của vùng gộp) và nội dung hiện có.
    ReDim targets(0 To cfg.MapCount - 1)
    ReDim texts(0 To cfg.MapCount - 1)
    For i = 0 To cfg.MapCount - 1
        Set ws = FindSheet(wb, cfg.Maps(i).SheetName)
        If ws Is Nothing Then
            message = "Không có sheet """ & cfg.Maps(i).SheetName & """ (ô đích " & _
                FormatTarget(cfg.Maps(i).SheetName, cfg.Maps(i).Address) & ") trong file. Biểu mẫu chưa được ghi."
            Exit Function
        End If
        Set targets(i) = ws.Range(cfg.Maps(i).Address).MergeArea.Cells(1, 1)
        raw = targets(i).Value2
        If VarType(raw) = vbString Then texts(i) = raw
    Next i

    key = MappingKey(maps, cfg.MapCount)
    If cfg.ConfirmedKey <> key Then
        For i = 0 To cfg.MapCount - 1
            raw = targets(i).Value2
            isPattern = (cfg.Maps(i).Column = cfg.DateColumn) And IsDatePattern(texts(i))
            If targets(i).HasFormula Or (Not IsEmpty(raw) And Not (VarType(raw) = vbString And Len(texts(i)) = 0) And Not isPattern) Then
                riskyCount = riskyCount + 1
                If riskyCount <= MAX_RISKY_SHOWN Then
                    risky = risky & vbLf & "  • " & FormatTarget(cfg.Maps(i).SheetName, cfg.Maps(i).Address)
                    If targets(i).HasFormula Then risky = risky & " — có công thức" Else risky = risky & " — có dữ liệu"
                End If
            End If
        Next i
        If riskyCount > MAX_RISKY_SHOWN Then risky = risky & vbLf & "  … và " & (riskyCount - MAX_RISKY_SHOWN) & " ô khác"
        If riskyCount > 0 Then
            If ShowMessage("Ghi dữ liệu ngày " & FormatDay(days(res.TargetIndex).DayNum) & _
                " lên biểu mẫu sẽ thay nội dung các ô sau:" & risky & vbLf & vbLf & _
                "Bấm OK để ghi đè (chỉ hỏi một lần cho bộ ô đã gán), Cancel để hủy.", _
                vbOKCancel + vbQuestion + vbDefaultButton2) <> vbOK Then
                message = ""
                Exit Function
            End If
        End If
    End If

    valueRow = days(res.TargetIndex).ValueRow
    c0 = LBound(values, 2)
    For i = 0 To cfg.MapCount - 1
        targets(i).Value = WriteValue(values(valueRow, c0 + cfg.Maps(i).Column), _
            cfg.Maps(i).Column = cfg.DateColumn, texts(i), wb.Date1904)
    Next i

    cfg.SourceSheet = source.Worksheet.Name
    cfg.SourceAddress = source.Address(False, False)
    cfg.CurrentDay = days(res.TargetIndex).DayNum
    cfg.ConfirmedKey = key
    SaveConfig wb, cfg
    message = res.Notice
    NavigateDay = True
End Function

' Dùng chung cho NavigateDay và LoadPrintDays: đọc cấu hình, vùng nguồn (khớp lại cột theo tiêu đề), kiểm tra ánh xạ,
' đọc giá trị mới nhất và dựng danh sách ngày. False -> message/style cho biết lý do.
Private Function PrepareSourceDays(ByVal wb As Workbook, ByRef cfg As NKTC_Config, ByRef source As Range, _
                                   ByRef values As Variant, ByRef days() As NKTC_SourceDay, ByRef dayCount As Long, _
                                   ByRef message As String, ByRef style As VbMsgBoxStyle) As Boolean
    Dim msg As String, headers() As String, errors As Collection, v As Variant, maps() As NKTC_Mapping
    Dim issues As Collection

    message = "": style = vbInformation
    If Not LoadConfig(wb, cfg) Then message = NotAssignedMessage(): Exit Function
    If cfg.MapCount = 0 Then message = NotAssignedMessage(): Exit Function
    style = vbExclamation

    Set source = ConfigSource(wb, cfg, msg)
    If source Is Nothing Then message = msg: Exit Function
    headers = HeaderTexts(source)
    msg = RemapConfig(cfg, headers, UBound(headers) + 1, RefText(source))
    If Len(msg) > 0 Then message = msg: Exit Function
    GetMaps cfg, maps
    Set errors = ValidateMappings(maps, cfg.MapCount, cfg.HeaderCount)
    If errors.Count > 0 Then
        For Each v In errors
            message = message & v & " "
        Next v
        message = message & "Bấm ""Xuất vào form"" để gán cột."
        Exit Function
    End If

    values = RangeGrid(source)
    BuildDayIndex values, cfg.DateColumn, source.Row, wb.Date1904, days, dayCount, issues
    PrepareSourceDays = True
End Function

' Hiển thị một ngày cụ thể trên form (cùng cơ chế, cùng hỏi xác nhận ghi đè như NavigateDay).
Public Function ShowDayOnForm(ByVal wb As Workbook, ByVal dayNum As Long, ByRef message As String, _
                              ByRef style As VbMsgBoxStyle) As Boolean
    ShowDayOnForm = NavigateDay(wb, CMD_GOTO, message, style, dayNum)
End Function

' Danh sách ngày của vùng nguồn cho chức năng In nhật ký. currentDay: ngày đang hiển thị (0 nếu chưa);
' formSheet: sheet của ô đích đầu tiên (Nothing nếu sheet không còn).
Public Function LoadPrintDays(ByVal wb As Workbook, ByRef days() As NKTC_SourceDay, ByRef dayCount As Long, _
                              ByRef currentDay As Long, ByRef formSheet As Worksheet, ByRef message As String) As Boolean
    Dim cfg As NKTC_Config, source As Range, values As Variant, style As VbMsgBoxStyle, maps() As NKTC_Mapping

    Set formSheet = Nothing
    currentDay = 0
    If Not PrepareSourceDays(wb, cfg, source, values, days, dayCount, message, style) Then Exit Function
    GetMaps cfg, maps
    currentDay = cfg.CurrentDay
    Set formSheet = FindSheet(wb, maps(0).SheetName)
    LoadPrintDays = True
End Function

'------------------------------------------------------------------------------
' Xuất vào form: hộp thoại chọn vùng nguồn, cột Ngày, các dòng gán "cột nguồn -> ô trên form"; Gán -> lưu cấu hình,
' biểu mẫu nhận dữ liệu ngày đầu tiên.
'------------------------------------------------------------------------------
Public Sub CmdXuatForm()
    Dim wb As Workbook, cfg As NKTC_Config, hasCfg As Boolean, f As frmXuatForm, initial As String
    Dim cols() As Long, targets() As String, n As Long, i As Long, accepted As Boolean, source As Range
    Dim newCfg As NKTC_Config, oldKey As String, errors As Collection, message As String, style As VbMsgBoxStyle
    Dim key As String, headers() As String, maps() As NKTC_Mapping, mapCount As Long

    If Not NeedWorkbook() Then Exit Sub
    Set wb = ActiveWorkbook
    hasCfg = LoadConfig(wb, cfg)
    initial = DefaultSourceText(wb, cfg, hasCfg)

    n = 0
    ReDim cols(0 To 0): ReDim targets(0 To 0)
    If hasCfg Then
        n = cfg.MapCount
        If n > 0 Then
            ReDim cols(0 To n - 1): ReDim targets(0 To n - 1)
            For i = 0 To n - 1
                cols(i) = cfg.Maps(i).Column
                targets(i) = FormatReference(cfg.Maps(i).SheetName, cfg.Maps(i).Address)
            Next i
        End If
        oldKey = cfg.ConfirmedKey
    End If

    Set f = New frmXuatForm
    If hasCfg Then
        GetHeaders cfg, headers
        f.Init initial, cfg.DateColumn, cols, targets, n, headers, cfg.HeaderCount
    Else
        ReDim headers(0 To 0)
        f.Init initial, -1, cols, targets, n, headers, 0
    End If
    Do
        f.Show
        If Len(f.PickRequest) = 0 Then Exit Do
        If f.PickRequest = "Nguon" Then
            f.ApplyPick PickRange("Xuất vào form", "Chọn vùng nguồn (nhật ký theo ngày, dòng đầu là tiêu đề). " & _
                "Vùng tự mở rộng tới hết dữ liệu của sheet.", f.PickCurrent)
        Else
            f.ApplyPick PickRange("Xuất vào form", "Chọn ô trên biểu mẫu sẽ nhận dữ liệu của cột.", f.PickCurrent, True)
        End If
    Loop
    accepted = f.Accepted
    If accepted Then
        Set source = f.Source
        newCfg.DateColumn = f.DateColumn
        n = f.RowCount
        If n > 0 Then
            ReDim cols(0 To n - 1): ReDim targets(0 To n - 1)
            For i = 0 To n - 1
                cols(i) = f.RowColumn(i)
                targets(i) = f.RowTarget(i)
            Next i
        End If
    End If
    Unload f
    Set f = Nothing
    If Not accepted Then Exit Sub

    newCfg.SourceSheet = source.Worksheet.Name
    newCfg.SourceAddress = source.Address(False, False)
    headers = HeaderTexts(source)
    SetHeaders newCfg, headers, UBound(headers) + 1
    Set errors = BuildMappingsFromRows(cols, targets, n, newCfg.HeaderCount, FormSheetName(), maps, mapCount)
    If errors.Count > 0 Then ShowInfo errors(1), vbExclamation: Exit Sub
    SetMaps newCfg, maps, mapCount
    ' Xác nhận ghi đè gắn với bộ ô đích (R2.5): cột nguồn đổi vị trí không làm mất xác nhận.
    key = MappingKey(maps, mapCount)
    If Len(oldKey) > 0 And oldKey = key Then newCfg.ConfirmedKey = oldKey
    SaveConfig wb, newCfg

    If NavigateDay(wb, CMD_FIRST, message, style) Then
        ShowInfo "Đã gán " & newCfg.MapCount & " ô; biểu mẫu đang hiển thị dữ liệu ngày đầu tiên." & vbLf & vbLf & _
            "Dùng các nút Ngày đầu / Ngày trước / Ngày sau / Ngày cuối (nhóm Chuyển ngày) để chuyển ngày." & _
            IIf(Len(message) > 0, vbLf & vbLf & message, "")
    ElseIf Len(message) > 0 Then
        ShowInfo message, style
    End If
End Sub

' Vùng nguồn gợi ý: cấu hình đã lưu -> sheet NK_TheoNgay đang mở hoặc mới nhất -> vùng đang chọn.
Private Function DefaultSourceText(ByVal wb As Workbook, ByRef cfg As NKTC_Config, ByVal hasCfg As Boolean) As String
    Dim r As Range, msg As String, ws As Worksheet, daily As Worksheet
    If hasCfg Then
        Set r = ConfigSource(wb, cfg, msg)
        If Not r Is Nothing Then DefaultSourceText = RefText(r): Exit Function
    End If
    If TypeName(ActiveSheet) = "Worksheet" Then
        If IsDailySheet(ActiveSheet.Name) Then Set daily = ActiveSheet
    End If
    If daily Is Nothing Then
        For Each ws In wb.Worksheets
            If IsDailySheet(ws.Name) Then Set daily = ws
        Next ws
    End If
    If Not daily Is Nothing Then
        DefaultSourceText = RefText(ExpandSource(daily.UsedRange))
    Else
        Set r = SelectedRange()
        If Not r Is Nothing Then DefaultSourceText = RefText(ExpandSource(r))
    End If
End Function

' NK_TheoNgay, NK_TheoNgay_2, … (không phân biệt hoa/thường).
Private Function IsDailySheet(ByVal name As String) As Boolean
    Dim base As String, rest As String, i As Long
    base = LCase$(DAILY_SHEET_BASE_NAME)
    name = LCase$(name)
    If name = base Then IsDailySheet = True: Exit Function
    If Left$(name, Len(base) + 1) <> base & "_" Then Exit Function
    rest = Mid$(name, Len(base) + 2)
    If Len(rest) = 0 Then Exit Function
    For i = 1 To Len(rest)
        If Mid$(rest, i, 1) < "0" Or Mid$(rest, i, 1) > "9" Then Exit Function
    Next i
    IsDailySheet = True
End Function

'------------------------------------------------------------------------------
' Hướng dẫn sử dụng.
'------------------------------------------------------------------------------
Public Sub ShowHelp()
    Dim printHelp As String
    printHelp = "3. In nhật ký (nhóm In)" & vbLf & _
        "   • ""In nhật ký"" → chọn Vùng in (mặc định là vùng in của form; nút ""Chọn"" để dùng vùng khác, kể cả trên sheet khác)," & vbLf & _
        "     Ngày in (đang hiển thị / một ngày / khoảng ngày / tất cả) và Đầu ra (máy in hoặc một file PDF gộp)." & vbLf & _
        "   • In nhiều hơn 50 ngày sẽ hỏi xác nhận; bấm Esc để dừng. Sau khi in, form trở về ngày đang xem."
    ShowInfo "Mọi chức năng nằm trên tab ""Nhật ký TC"" của thanh Ribbon." & vbLf & vbLf & _
        "0. Tạo bảng nhập tiến độ (nhóm Tiến độ)" & vbLf & _
        "   • Bấm ""Tạo bảng"" → chọn sheet mới hoặc sheet hiện tại (ô bắt đầu), số dòng → ""Tạo bảng""." & vbLf & _
        "   • Điền Nội dung công việc, ngày Bắt đầu, ngày Kết thúc (dd/mm/yyyy)." & vbLf & vbLf & _
        "1. Xuất nhật ký theo ngày (nhóm Tiến độ)" & vbLf & _
        "   • Đặt con trỏ trong bảng tiến độ (hoặc bôi chọn bảng) → ""Xuất theo ngày""." & vbLf & _
        "   • Kiểm tra Vùng dữ liệu (nút ""Chọn"" để chọn lại trên sheet) và 3 cột → ""Xuất nhật ký""." & vbLf & _
        "   • Add-in tạo sheet mới NK_TheoNgay (không ghi đè sheet có sẵn)." & vbLf & vbLf & _
        "2. Xuất theo form mẫu (nhóm Form mẫu và Chuyển ngày)" & vbLf & _
        "   • Trong NK_TheoNgay có thể thêm cột: Thời tiết, Số công nhân, Nội dung nghiệm thu…" & vbLf & _
        "   • ""Chèn form mẫu"" thêm sheet ""Mẫu NKTC""." & vbLf & _
        "   • ""Xuất vào form"": kiểm tra Vùng nguồn (dòng đầu là tiêu đề, không có ô gộp) và cột Ngày;" & vbLf & _
        "     chọn Cột nguồn + Ô trên form (nút ""Chọn"") → ""Thêm""; lặp lại cho từng cột → ""Gán""." & vbLf & _
        "   • Cột đã gán được nhớ theo tên tiêu đề: thêm/chèn cột, thêm ngày không cần gán lại." & vbLf & _
        "   • Sau khi gán: Ngày đầu / Ngày trước / Ngày sau / Ngày cuối để chuyển ngày. Gán cột Ngày vào" & vbLf & _
        "     dòng ""Ngày … tháng … năm …"" của biểu mẫu: add-in tự điền ngày, tháng, năm." & vbLf & vbLf & _
        printHelp
End Sub
