Attribute VB_Name = "modNKTC_FormCommands"
'==============================================================================
' NHAT KY TC ? F2: Chen form mau, Xuat vao form (gan cot), Chuyen ngay, Huong dan.
' Tuong ung src/commands/{formCommands,templateCommands}.ts va src/excel/{settingsStore,templateWriter}.ts.
' Chi ghi vao cac o dich da gan (CLAUDE.md ?4); o dang co cong thuc/du lieu -> hoi truoc khi ghi de (R2.5).
'==============================================================================
Option Explicit

Private Const CONFIG_NS As String = "urn:nhatkytc:config:1"
Private Const MAX_RISKY_SHOWN As Long = 15

' Ten sheet bieu mau (co san trong add-in, lay tu templates/Mau_NKTC.xlsx khi build).
Public Function FormSheetName() As String
    FormSheetName = U("M\u1eabu NKTC")
End Function

Private Function NotAssignedMessage() As String
    NotAssignedMessage = U("Ch\u01b0a g\u00e1n c\u1ed9t v\u00e0o form. B\u1ea5m ""Xu\u1ea5t v\u00e0o form"" (nh\u00f3m Form m\u1eabu) \u0111\u1ec3 ch\u1ecdn v\u00f9ng ngu\u1ed3n v\u00e0 g\u00e1n c\u1ed9t ") & _
        U("tr\u01b0\u1edbc \u2014 chuy\u1ec3n ng\u00e0y ch\u1ec9 ho\u1ea1t \u0111\u1ed9ng sau khi \u0111\u00e3 g\u00e1n.")
End Function

Private Function NeedWorkbook() As Boolean
    If ActiveWorkbook Is Nothing Then
        ShowInfo U("H\u00e3y m\u1edf m\u1ed9t file Excel tr\u01b0\u1edbc."), vbExclamation
    Else
        NeedWorkbook = True
    End If
End Function

'------------------------------------------------------------------------------
' Cau hinh luu trong file (CustomXMLPart) ? giu lai khi nguoi dung luu file Excel (R2.8).
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

' Vung nguon cua cau hinh, da tu mo rong toi het du lieu cua sheet. msg: ly do neu Nothing.
Private Function ConfigSource(ByVal wb As Workbook, ByRef cfg As NKTC_Config, ByRef msg As String) As Range
    Dim ws As Worksheet, r As Range
    Set ws = FindSheet(wb, cfg.SourceSheet)
    If ws Is Nothing Then
        msg = U("Kh\u00f4ng c\u00f3 sheet """) & cfg.SourceSheet & U(""" (v\u00f9ng ngu\u1ed3n \u0111\u00e3 g\u00e1n) trong file. B\u1ea5m ""Xu\u1ea5t v\u00e0o form"" \u0111\u1ec3 ch\u1ecdn l\u1ea1i v\u00f9ng ngu\u1ed3n.")
        Exit Function
    End If
    On Error Resume Next
    Set r = ws.Range(cfg.SourceAddress)
    On Error GoTo 0
    If r Is Nothing Then msg = U("V\u00f9ng ngu\u1ed3n \u0111\u00e3 l\u01b0u kh\u00f4ng h\u1ee3p l\u1ec7. B\u1ea5m ""Xu\u1ea5t v\u00e0o form"" \u0111\u1ec3 ch\u1ecdn l\u1ea1i."): Exit Function
    Set ConfigSource = ExpandSource(r)
End Function

'------------------------------------------------------------------------------
' Chen form mau: chep sheet "Mau NKTC" tu add-in vao cuoi file dang mo. Khong ghi de sheet co san.
'------------------------------------------------------------------------------
Public Sub CmdChenForm()
    Dim wb As Workbook, ws As Worksheet
    If Not NeedWorkbook() Then Exit Sub
    Set wb = ActiveWorkbook
    If Not FindSheet(wb, FormSheetName()) Is Nothing Then
        ShowInfo U("File \u0111\u00e3 c\u00f3 sheet """) & FormSheetName() & U(""". Add-in kh\u00f4ng ghi \u0111\u00e8 \u2014 \u0111\u1ed5i t\u00ean ho\u1eb7c x\u00f3a sheet \u0111\u00f3 n\u1ebfu mu\u1ed1n ch\u00e8n m\u1eabu m\u1edbi."), vbExclamation
        Exit Sub
    End If
    If wb.ProtectStructure Then
        ShowInfo U("File \u0111ang kh\u00f3a c\u1ea5u tr\u00fac (Protect Workbook) n\u00ean kh\u00f4ng th\u00eam \u0111\u01b0\u1ee3c sheet."), vbExclamation
        Exit Sub
    End If
    ThisWorkbook.Worksheets(FormSheetName()).Copy After:=wb.Sheets(wb.Sheets.Count)
    Set ws = FindSheet(wb, FormSheetName())
    If Not ws Is Nothing Then ws.Activate
    ShowInfo U("\u0110\u00e3 ch\u00e8n sheet """) & FormSheetName() & """." & vbLf & vbLf & _
        U("Ti\u1ebfp theo: b\u1ea5m ""Xu\u1ea5t v\u00e0o form"" \u0111\u1ec3 ch\u1ecdn v\u00f9ng nh\u1eadt k\u00fd theo ng\u00e0y v\u00e0 g\u00e1n c\u1ed9t v\u00e0o c\u00e1c \u00f4 c\u1ee7a bi\u1ec3u m\u1eabu.")
End Sub

'------------------------------------------------------------------------------
' Chuyen ngay (R2.3?R2.7, R2.15, R2.17).
'------------------------------------------------------------------------------
Public Sub CmdChuyenNgay(ByVal command As Long)
    Dim message As String, style As VbMsgBoxStyle
    If Not NeedWorkbook() Then Exit Sub
    NavigateDay ActiveWorkbook, command, message, style
    If Len(message) > 0 Then ShowInfo message, style
End Sub

' Ghi du lieu mot ngay len form: doc lai nguon (moi nhat), khop cot theo ten tieu de, chon ngay, hoi truoc khi
' ghi de o co cong thuc/du lieu (mot lan cho moi bo o dich), roi ghi. True neu da ghi; message: thong bao kem theo.
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

    ' O dich (o gop -> o dau cua vung gop) va noi dung hien co.
    ReDim targets(0 To cfg.MapCount - 1)
    ReDim texts(0 To cfg.MapCount - 1)
    For i = 0 To cfg.MapCount - 1
        Set ws = FindSheet(wb, cfg.Maps(i).SheetName)
        If ws Is Nothing Then
            message = U("Kh\u00f4ng c\u00f3 sheet """) & cfg.Maps(i).SheetName & U(""" (\u00f4 \u0111\u00edch ") & _
                FormatTarget(cfg.Maps(i).SheetName, cfg.Maps(i).Address) & U(") trong file. Bi\u1ec3u m\u1eabu ch\u01b0a \u0111\u01b0\u1ee3c ghi.")
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
                    risky = risky & vbLf & U("  \u2022 ") & FormatTarget(cfg.Maps(i).SheetName, cfg.Maps(i).Address)
                    If targets(i).HasFormula Then risky = risky & U(" \u2014 c\u00f3 c\u00f4ng th\u1ee9c") Else risky = risky & U(" \u2014 c\u00f3 d\u1eef li\u1ec7u")
                End If
            End If
        Next i
        If riskyCount > MAX_RISKY_SHOWN Then risky = risky & vbLf & U("  \u2026 v\u00e0 ") & (riskyCount - MAX_RISKY_SHOWN) & U(" \u00f4 kh\u00e1c")
        If riskyCount > 0 Then
            If ShowMessage(U("Ghi d\u1eef li\u1ec7u ng\u00e0y ") & FormatDay(days(res.TargetIndex).DayNum) & _
                U(" l\u00ean bi\u1ec3u m\u1eabu s\u1ebd thay n\u1ed9i dung c\u00e1c \u00f4 sau:") & risky & vbLf & vbLf & _
                U("B\u1ea5m OK \u0111\u1ec3 ghi \u0111\u00e8 (ch\u1ec9 h\u1ecfi m\u1ed9t l\u1ea7n cho b\u1ed9 \u00f4 \u0111\u00e3 g\u00e1n), Cancel \u0111\u1ec3 h\u1ee7y."), _
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

' Dung chung cho NavigateDay va LoadPrintDays: doc cau hinh, vung nguon (khop lai cot theo tieu de), kiem tra anh xa,
' doc gia tri moi nhat va dung danh sach ngay. False -> message/style cho biet ly do.
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
        message = message & U("B\u1ea5m ""Xu\u1ea5t v\u00e0o form"" \u0111\u1ec3 g\u00e1n c\u1ed9t.")
        Exit Function
    End If

    values = RangeGrid(source)
    BuildDayIndex values, cfg.DateColumn, source.Row, wb.Date1904, days, dayCount, issues
    PrepareSourceDays = True
End Function

' Hien thi mot ngay cu the tren form (cung co che, cung hoi xac nhan ghi de nhu NavigateDay).
Public Function ShowDayOnForm(ByVal wb As Workbook, ByVal dayNum As Long, ByRef message As String, _
                              ByRef style As VbMsgBoxStyle) As Boolean
    ShowDayOnForm = NavigateDay(wb, CMD_GOTO, message, style, dayNum)
End Function

' Danh sach ngay cua vung nguon cho chuc nang In nhat ky. currentDay: ngay dang hien thi (0 neu chua);
' formSheet: sheet cua o dich dau tien (Nothing neu sheet khong con).
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
' Xuat vao form: hop thoai chon vung nguon, cot Ngay, cac dong gan "cot nguon -> o tren form"; Gan -> luu cau hinh,
' bieu mau nhan du lieu ngay dau tien.
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
            f.ApplyPick PickRange(U("Xu\u1ea5t v\u00e0o form"), U("Ch\u1ecdn v\u00f9ng ngu\u1ed3n (nh\u1eadt k\u00fd theo ng\u00e0y, d\u00f2ng \u0111\u1ea7u l\u00e0 ti\u00eau \u0111\u1ec1). ") & _
                U("V\u00f9ng t\u1ef1 m\u1edf r\u1ed9ng t\u1edbi h\u1ebft d\u1eef li\u1ec7u c\u1ee7a sheet."), f.PickCurrent)
        Else
            f.ApplyPick PickRange(U("Xu\u1ea5t v\u00e0o form"), U("Ch\u1ecdn \u00f4 tr\u00ean bi\u1ec3u m\u1eabu s\u1ebd nh\u1eadn d\u1eef li\u1ec7u c\u1ee7a c\u1ed9t."), f.PickCurrent, True)
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
    ' Xac nhan ghi de gan voi bo o dich (R2.5): cot nguon doi vi tri khong lam mat xac nhan.
    key = MappingKey(maps, mapCount)
    If Len(oldKey) > 0 And oldKey = key Then newCfg.ConfirmedKey = oldKey
    SaveConfig wb, newCfg

    If NavigateDay(wb, CMD_FIRST, message, style) Then
        ShowInfo U("\u0110\u00e3 g\u00e1n ") & newCfg.MapCount & U(" \u00f4; bi\u1ec3u m\u1eabu \u0111ang hi\u1ec3n th\u1ecb d\u1eef li\u1ec7u ng\u00e0y \u0111\u1ea7u ti\u00ean.") & vbLf & vbLf & _
            U("D\u00f9ng c\u00e1c n\u00fat Ng\u00e0y \u0111\u1ea7u / Ng\u00e0y tr\u01b0\u1edbc / Ng\u00e0y sau / Ng\u00e0y cu\u1ed1i (nh\u00f3m Chuy\u1ec3n ng\u00e0y) \u0111\u1ec3 chuy\u1ec3n ng\u00e0y.") & _
            IIf(Len(message) > 0, vbLf & vbLf & message, "")
    ElseIf Len(message) > 0 Then
        ShowInfo message, style
    End If
End Sub

' Vung nguon goi y: cau hinh da luu -> sheet NK_TheoNgay dang mo hoac moi nhat -> vung dang chon.
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

' NK_TheoNgay, NK_TheoNgay_2, ? (khong phan biet hoa/thuong).
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
' Huong dan su dung.
'------------------------------------------------------------------------------
Public Sub ShowHelp()
    Dim printHelp As String
    printHelp = U("3. In nh\u1eadt k\u00fd (nh\u00f3m In)") & vbLf & _
        U("   \u2022 ""In nh\u1eadt k\u00fd"" \u2192 ch\u1ecdn V\u00f9ng in (m\u1eb7c \u0111\u1ecbnh l\u00e0 v\u00f9ng in c\u1ee7a form; n\u00fat ""Ch\u1ecdn"" \u0111\u1ec3 d\u00f9ng v\u00f9ng kh\u00e1c, k\u1ec3 c\u1ea3 tr\u00ean sheet kh\u00e1c),") & vbLf & _
        U("     Ng\u00e0y in (\u0111ang hi\u1ec3n th\u1ecb / m\u1ed9t ng\u00e0y / kho\u1ea3ng ng\u00e0y / t\u1ea5t c\u1ea3) v\u00e0 \u0110\u1ea7u ra (m\u00e1y in ho\u1eb7c m\u1ed9t file PDF g\u1ed9p).") & vbLf & _
        U("   \u2022 In nhi\u1ec1u h\u01a1n 50 ng\u00e0y s\u1ebd h\u1ecfi x\u00e1c nh\u1eadn; b\u1ea5m Esc \u0111\u1ec3 d\u1eebng. Sau khi in, form tr\u1edf v\u1ec1 ng\u00e0y \u0111ang xem.")
    ShowInfo U("M\u1ecdi ch\u1ee9c n\u0103ng n\u1eb1m tr\u00ean tab ""Nh\u1eadt k\u00fd TC"" c\u1ee7a thanh Ribbon.") & vbLf & vbLf & _
        U("0. T\u1ea1o b\u1ea3ng nh\u1eadp ti\u1ebfn \u0111\u1ed9 (nh\u00f3m Ti\u1ebfn \u0111\u1ed9)") & vbLf & _
        U("   \u2022 B\u1ea5m ""T\u1ea1o b\u1ea3ng"" \u2192 ch\u1ecdn sheet m\u1edbi ho\u1eb7c sheet hi\u1ec7n t\u1ea1i (\u00f4 b\u1eaft \u0111\u1ea7u), s\u1ed1 d\u00f2ng \u2192 ""T\u1ea1o b\u1ea3ng"".") & vbLf & _
        U("   \u2022 \u0110i\u1ec1n N\u1ed9i dung c\u00f4ng vi\u1ec7c, ng\u00e0y B\u1eaft \u0111\u1ea7u, ng\u00e0y K\u1ebft th\u00fac (dd/mm/yyyy).") & vbLf & vbLf & _
        U("1. Xu\u1ea5t nh\u1eadt k\u00fd theo ng\u00e0y (nh\u00f3m Ti\u1ebfn \u0111\u1ed9)") & vbLf & _
        U("   \u2022 \u0110\u1eb7t con tr\u1ecf trong b\u1ea3ng ti\u1ebfn \u0111\u1ed9 (ho\u1eb7c b\u00f4i ch\u1ecdn b\u1ea3ng) \u2192 ""Xu\u1ea5t theo ng\u00e0y"".") & vbLf & _
        U("   \u2022 Ki\u1ec3m tra V\u00f9ng d\u1eef li\u1ec7u (n\u00fat ""Ch\u1ecdn"" \u0111\u1ec3 ch\u1ecdn l\u1ea1i tr\u00ean sheet) v\u00e0 3 c\u1ed9t \u2192 ""Xu\u1ea5t nh\u1eadt k\u00fd"".") & vbLf & _
        U("   \u2022 Add-in t\u1ea1o sheet m\u1edbi NK_TheoNgay (kh\u00f4ng ghi \u0111\u00e8 sheet c\u00f3 s\u1eb5n).") & vbLf & vbLf & _
        U("2. Xu\u1ea5t theo form m\u1eabu (nh\u00f3m Form m\u1eabu v\u00e0 Chuy\u1ec3n ng\u00e0y)") & vbLf & _
        U("   \u2022 Trong NK_TheoNgay c\u00f3 th\u1ec3 th\u00eam c\u1ed9t: Th\u1eddi ti\u1ebft, S\u1ed1 c\u00f4ng nh\u00e2n, N\u1ed9i dung nghi\u1ec7m thu\u2026") & vbLf & _
        U("   \u2022 ""Ch\u00e8n form m\u1eabu"" th\u00eam sheet ""M\u1eabu NKTC"".") & vbLf & _
        U("   \u2022 ""Xu\u1ea5t v\u00e0o form"": ki\u1ec3m tra V\u00f9ng ngu\u1ed3n (d\u00f2ng \u0111\u1ea7u l\u00e0 ti\u00eau \u0111\u1ec1, kh\u00f4ng c\u00f3 \u00f4 g\u1ed9p) v\u00e0 c\u1ed9t Ng\u00e0y;") & vbLf & _
        U("     ch\u1ecdn C\u1ed9t ngu\u1ed3n + \u00d4 tr\u00ean form (n\u00fat ""Ch\u1ecdn"") \u2192 ""Th\u00eam""; l\u1eb7p l\u1ea1i cho t\u1eebng c\u1ed9t \u2192 ""G\u00e1n"".") & vbLf & _
        U("   \u2022 C\u1ed9t \u0111\u00e3 g\u00e1n \u0111\u01b0\u1ee3c nh\u1edb theo t\u00ean ti\u00eau \u0111\u1ec1: th\u00eam/ch\u00e8n c\u1ed9t, th\u00eam ng\u00e0y kh\u00f4ng c\u1ea7n g\u00e1n l\u1ea1i.") & vbLf & _
        U("   \u2022 Sau khi g\u00e1n: Ng\u00e0y \u0111\u1ea7u / Ng\u00e0y tr\u01b0\u1edbc / Ng\u00e0y sau / Ng\u00e0y cu\u1ed1i \u0111\u1ec3 chuy\u1ec3n ng\u00e0y. G\u00e1n c\u1ed9t Ng\u00e0y v\u00e0o") & vbLf & _
        U("     d\u00f2ng ""Ng\u00e0y \u2026 th\u00e1ng \u2026 n\u0103m \u2026"" c\u1ee7a bi\u1ec3u m\u1eabu: add-in t\u1ef1 \u0111i\u1ec1n ng\u00e0y, th\u00e1ng, n\u0103m.") & vbLf & vbLf & _
        printHelp
End Sub
