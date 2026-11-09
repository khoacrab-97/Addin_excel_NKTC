Attribute VB_Name = "modNKTC_Config"
'==============================================================================
' NHẬT KÝ TC – F2 Xuất vào form: địa chỉ ô đích, dòng gán, cấu hình lưu trong file (không phụ thuộc Excel).
' Tương ứng src/core/address/reference.ts, src/core/mapping/{formRows,config,headerMatch}.ts của bản Office.js.
'==============================================================================
Option Explicit

Public Const MAX_FORM_ROWS As Long = 500
Private Const MAX_SHEET_NAME As Long = 31
Private Const MAX_COLUMN As Long = 16384
Private Const MAX_ROW As Long = 1048576
Private Const CONFIG_HEADER As String = "NKTC-CONFIG"
Private Const CONFIG_VERSION As String = "1"

' Cấu hình gán đã lưu (R2.8): vùng nguồn, dòng tiêu đề lúc gán, cột Ngày, các gán, ngày đang hiển thị.
Public Type NKTC_Config
    SourceSheet As String
    SourceAddress As String    ' không có $, ví dụ "A1:F40"
    DateColumn As Long         ' 0-based trong vùng nguồn
    CurrentDay As Long         ' 0 = chưa hiển thị ngày nào
    ConfirmedKey As String     ' bộ ô đích đã được xác nhận ghi đè (R2.5)
    HeaderCount As Long
    Headers() As String        ' 0-based
    MapCount As Long
    Maps() As NKTC_Mapping     ' 0-based
End Type

'------------------------------------------------------------------------------
' Địa chỉ ô đích: "Mẫu NKTC!E2", "'Mẫu NKTC'!$E$2", "E2" (không ghi sheet).
'------------------------------------------------------------------------------
Private Function ValidSheetName(ByVal name As String) As Boolean
    Dim i As Long
    If Len(name) = 0 Or Len(name) > MAX_SHEET_NAME Then Exit Function
    For i = 1 To Len(name)
        If InStr(1, "[]:*?/\", Mid$(name, i, 1), vbBinaryCompare) > 0 Then Exit Function
    Next i
    ValidSheetName = True
End Function

' Một ô: 1–3 chữ cái + số dòng, trong giới hạn sheet. Trả về địa chỉ viết hoa không $; rỗng nếu không hợp lệ.
Public Function NormalizeCellAddress(ByVal address As String) As String
    Dim i As Long, letters As String, digits As String, ch As String, col As Long, k As Long
    address = UCase$(Replace(Trim$(address), "$", ""))
    For i = 1 To Len(address)
        ch = Mid$(address, i, 1)
        If ch >= "A" And ch <= "Z" Then
            If Len(digits) > 0 Then Exit Function
            letters = letters & ch
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            Exit Function
        End If
    Next i
    If Len(letters) = 0 Or Len(letters) > 3 Or Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
    For k = 1 To Len(letters)
        col = col * 26 + (Asc(Mid$(letters, k, 1)) - 64)
    Next k
    If col > MAX_COLUMN Or CLng(digits) < 1 Or CLng(digits) > MAX_ROW Then Exit Function
    NormalizeCellAddress = letters & CLng(digits)
End Function

' Đọc tham chiếu một ô. sheetName = "" nếu người dùng không ghi tên sheet.
Public Function ParseCellReference(ByVal text As String, ByRef sheetName As String, ByRef address As String) As Boolean
    Dim sep As Long, rawSheet As String
    sheetName = "": address = ""
    text = Trim$(text)
    If Len(text) = 0 Or Len(text) > 80 Then Exit Function
    sep = InStrRev(text, "!")
    address = NormalizeCellAddress(Mid$(text, sep + 1))
    If Len(address) = 0 Then Exit Function
    If sep > 0 Then
        rawSheet = Trim$(Left$(text, sep - 1))
        If Len(rawSheet) >= 2 And Left$(rawSheet, 1) = "'" And Right$(rawSheet, 1) = "'" Then
            rawSheet = Replace(Mid$(rawSheet, 2, Len(rawSheet) - 2), "''", "'")
        End If
        If Not ValidSheetName(rawSheet) Then address = "": Exit Function
        sheetName = rawSheet
    End If
    ParseCellReference = True
End Function

' Hiển thị cho người dùng: "Mẫu NKTC!E2".
Public Function FormatReference(ByVal sheetName As String, ByVal address As String) As String
    FormatReference = sheetName & "!" & address
End Function

' Dạng Excel hiểu được (Application.Range): "'Mẫu NKTC'!E2".
Public Function QuotedReference(ByVal sheetName As String, ByVal address As String) As String
    QuotedReference = "'" & Replace(sheetName, "'", "''") & "'!" & address
End Function

'------------------------------------------------------------------------------
' Dòng gán của hộp thoại (R2.1, R2.2, R2.12). cols(i) = -1: chưa chọn cột; targets(i): ô trên form.
' Dòng trống hoàn toàn bỏ qua; ô không ghi sheet thuộc defaultSheet; một ô chỉ nhận một cột.
' Trả về danh sách lỗi (rỗng nếu hợp lệ).
'------------------------------------------------------------------------------
Public Function BuildMappingsFromRows(ByRef cols() As Long, ByRef targets() As String, ByVal rowCount As Long, _
                                      ByVal columnCount As Long, ByVal defaultSheet As String, _
                                      ByRef maps() As NKTC_Mapping, ByRef mapCount As Long) As Collection
    Dim errors As New Collection, owners As New Collection, i As Long, label As String, text As String
    Dim sheet As String, address As String, key As String, owner As Variant
    mapCount = 0
    ReDim maps(0 To 0)
    If rowCount > MAX_FORM_ROWS Then errors.Add "Tối đa " & MAX_FORM_ROWS & " dòng gán.": Set BuildMappingsFromRows = errors: Exit Function
    For i = 0 To rowCount - 1
        label = "Dòng " & (i + 1)
        text = Trim$(targets(i))
        If cols(i) < 0 And Len(text) = 0 Then GoTo TiepTheo
        If cols(i) < 0 Or cols(i) >= columnCount Then
            errors.Add label & ": chưa chọn cột nguồn."
        ElseIf Len(text) = 0 Then
            errors.Add label & ": chưa chọn ô trên form."
        ElseIf Not ParseCellReference(text, sheet, address) Then
            errors.Add label & ": """ & text & """ không phải một ô hợp lệ (ví dụ Mẫu NKTC!E2)."
        Else
            If Len(sheet) = 0 Then sheet = defaultSheet
            key = TargetKey(sheet, address)
            owner = Empty
            On Error Resume Next
            owner = owners(key)
            On Error GoTo 0
            If Not IsEmpty(owner) Then
                errors.Add label & ": ô " & FormatReference(sheet, address) & " đã được gán ở dòng " & (owner + 1) & "."
            Else
                owners.Add i, key
                ReDim Preserve maps(0 To mapCount)
                maps(mapCount).Column = cols(i)
                maps(mapCount).SheetName = sheet
                maps(mapCount).Address = address
                mapCount = mapCount + 1
            End If
        End If
TiepTheo:
    Next i
    If errors.Count = 0 And mapCount = 0 Then errors.Add "Chưa có dòng gán nào. Chọn cột nguồn và ô trên form rồi bấm Thêm."
    If errors.Count > 0 Then mapCount = 0
    Set BuildMappingsFromRows = errors
End Function

' Cột có tiêu đề trống (thường do ô bị gộp — merge): "1, 3" (thứ tự 1-based); rỗng nếu không có.
Public Function BlankHeaderList(ByRef headers() As String, ByVal count As Long) As String
    Dim i As Long, out As String
    For i = 0 To count - 1
        If Len(Trim$(headers(i))) = 0 Then
            If Len(out) > 0 Then out = out & ", "
            out = out & (i + 1)
        End If
    Next i
    BlankHeaderList = out
End Function

'------------------------------------------------------------------------------
' Khớp lại cấu hình theo dòng tiêu đề mới (R2.17): trả về thông báo cột thiếu; rỗng nếu khớp đủ (cfg đã cập nhật).
'------------------------------------------------------------------------------
Public Function RemapConfig(ByRef cfg As NKTC_Config, ByRef newHeaders() As String, ByVal newCount As Long, _
                            ByVal sourceText As String) As String
    Dim pos() As Long, oldH() As String, newH() As String, i As Long, missing As String, cols() As Long, dateCol As Long
    ReDim oldH(0 To cfg.HeaderCount - 1)
    For i = 0 To cfg.HeaderCount - 1
        oldH(i) = cfg.Headers(i)
    Next i
    ReDim newH(0 To newCount - 1)
    For i = 0 To newCount - 1
        newH(i) = newHeaders(i)
    Next i
    pos = MatchHeaders(oldH, newH)

    dateCol = PositionOf(pos, cfg.DateColumn)
    If dateCol < 0 Then missing = "cột Ngày """ & HeaderOf(cfg, cfg.DateColumn) & """"
    If cfg.MapCount > 0 Then ReDim cols(0 To cfg.MapCount - 1)
    For i = 0 To cfg.MapCount - 1
        cols(i) = PositionOf(pos, cfg.Maps(i).Column)
        If cols(i) < 0 Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & "cột """ & HeaderOf(cfg, cfg.Maps(i).Column) & """ (gán vào " & _
                FormatTarget(cfg.Maps(i).SheetName, cfg.Maps(i).Address) & ")"
        End If
    Next i
    If Len(missing) > 0 Then
        RemapConfig = "Không tìm thấy " & missing & " trong dòng tiêu đề của vùng nguồn " & sourceText & _
            " — cột có thể đã bị xóa hoặc đổi tên tiêu đề. Biểu mẫu chưa được ghi. " & _
            "Bấm ""Xuất vào form"" để chọn lại cột cho các dòng này."
        Exit Function
    End If
    cfg.DateColumn = dateCol
    For i = 0 To cfg.MapCount - 1
        cfg.Maps(i).Column = cols(i)
    Next i
    SetHeaders cfg, newH, newCount
End Function

'------------------------------------------------------------------------------
' Chép mảng vào/ra cấu hình (không truyền thẳng mảng thành phần của Type vào thủ tục khác).
'------------------------------------------------------------------------------
Public Sub SetHeaders(ByRef cfg As NKTC_Config, ByRef headers() As String, ByVal count As Long)
    Dim i As Long
    cfg.HeaderCount = count
    If count = 0 Then ReDim cfg.Headers(0 To 0): Exit Sub
    ReDim cfg.Headers(0 To count - 1)
    For i = 0 To count - 1
        cfg.Headers(i) = headers(LBound(headers) + i)
    Next i
End Sub

Public Sub GetHeaders(ByRef cfg As NKTC_Config, ByRef headers() As String)
    Dim i As Long
    If cfg.HeaderCount = 0 Then ReDim headers(0 To 0): Exit Sub
    ReDim headers(0 To cfg.HeaderCount - 1)
    For i = 0 To cfg.HeaderCount - 1
        headers(i) = cfg.Headers(i)
    Next i
End Sub

Public Sub SetMaps(ByRef cfg As NKTC_Config, ByRef maps() As NKTC_Mapping, ByVal count As Long)
    Dim i As Long
    cfg.MapCount = count
    If count = 0 Then ReDim cfg.Maps(0 To 0): Exit Sub
    ReDim cfg.Maps(0 To count - 1)
    For i = 0 To count - 1
        cfg.Maps(i) = maps(i)
    Next i
End Sub

Public Sub GetMaps(ByRef cfg As NKTC_Config, ByRef maps() As NKTC_Mapping)
    Dim i As Long
    If cfg.MapCount = 0 Then ReDim maps(0 To 0): Exit Sub
    ReDim maps(0 To cfg.MapCount - 1)
    For i = 0 To cfg.MapCount - 1
        maps(i) = cfg.Maps(i)
    Next i
End Sub

Private Function PositionOf(ByRef pos() As Long, ByVal column As Long) As Long
    PositionOf = -1
    If column >= LBound(pos) And column <= UBound(pos) Then PositionOf = pos(column)
End Function

Private Function HeaderOf(ByRef cfg As NKTC_Config, ByVal column As Long) As String
    If column >= 0 And column < cfg.HeaderCount Then HeaderOf = cfg.Headers(column)
End Function

'------------------------------------------------------------------------------
' Lưu / đọc cấu hình dạng chữ: mỗi dòng một mục, các trường cách nhau bằng Tab.
'   NKTC-CONFIG <Tab> 1
'   source <Tab> tên sheet <Tab> A1:F40
'   date <Tab> 0 | current <Tab> 46188 | confirmed <Tab> khóa
'   header <Tab> tiêu đề (theo thứ tự cột)
'   map <Tab> cột <Tab> tên sheet <Tab> E2
'------------------------------------------------------------------------------
' Tab, xuống dòng, \ trong một trường -> \t, \n, \r, \\; ký tự điều khiển khác bỏ đi.
Public Function EscapeField(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If ch = "\" Then
            out = out & "\\"
        ElseIf code = 9 Then
            out = out & "\t"
        ElseIf code = 10 Then
            out = out & "\n"
        ElseIf code = 13 Then
            out = out & "\r"
        ElseIf code >= 0 And code < 32 Then
            ' bỏ
        Else
            out = out & ch
        End If
    Next i
    EscapeField = out
End Function

Public Function UnescapeField(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            i = i + 1
            Select Case Mid$(s, i, 1)
                Case "t": out = out & vbTab
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case Else: out = out & Mid$(s, i, 1)
            End Select
        Else
            out = out & ch
        End If
        i = i + 1
    Loop
    UnescapeField = out
End Function

' Chữ đặt trong phần tử XML (CustomXMLPart): & < > và Tab/xuống dòng thành tham chiếu ký tự.
Public Function XmlText(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbTab, "&#9;")
    s = Replace(s, vbLf, "&#10;")
    s = Replace(s, vbCr, "&#13;")
    XmlText = s
End Function

Public Function SerializeConfig(ByRef cfg As NKTC_Config) As String
    Dim lines As Collection, i As Long, out As String, v As Variant
    Set lines = New Collection
    lines.Add CONFIG_HEADER & vbTab & CONFIG_VERSION
    lines.Add "source" & vbTab & EscapeField(cfg.SourceSheet) & vbTab & cfg.SourceAddress
    lines.Add "date" & vbTab & cfg.DateColumn
    lines.Add "current" & vbTab & cfg.CurrentDay
    lines.Add "confirmed" & vbTab & EscapeField(cfg.ConfirmedKey)
    For i = 0 To cfg.HeaderCount - 1
        lines.Add "header" & vbTab & EscapeField(cfg.Headers(i))
    Next i
    For i = 0 To cfg.MapCount - 1
        lines.Add "map" & vbTab & cfg.Maps(i).Column & vbTab & EscapeField(cfg.Maps(i).SheetName) & vbTab & cfg.Maps(i).Address
    Next i
    For Each v In lines
        If Len(out) > 0 Then out = out & vbLf
        out = out & v
    Next v
    SerializeConfig = out
End Function

Private Function IsCount(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsCount = True
End Function

' Đọc cấu hình; False nếu không phải cấu hình hợp lệ (cfg khi đó không dùng được).
Public Function ParseConfig(ByVal text As String, ByRef cfg As NKTC_Config) As Boolean
    Dim lines As Variant, f As Variant, i As Long, sheet As String, address As String, empty0 As NKTC_Config
    cfg = empty0
    cfg.DateColumn = -1
    ReDim cfg.Headers(0 To 0)
    ReDim cfg.Maps(0 To 0)
    lines = Split(Replace(text, vbCr, ""), vbLf)
    If UBound(lines) < 0 Then Exit Function
    If lines(0) <> CONFIG_HEADER & vbTab & CONFIG_VERSION Then Exit Function
    For i = 1 To UBound(lines)
        f = Split(lines(i), vbTab)
        If UBound(f) >= 1 Then
            Select Case f(0)
                Case "source"
                    If UBound(f) < 2 Then Exit Function
                    cfg.SourceSheet = UnescapeField(f(1))
                    cfg.SourceAddress = f(2)
                Case "date"
                    If Not IsCount(f(1)) Then Exit Function
                    cfg.DateColumn = CLng(f(1))
                Case "current"
                    If Not IsCount(f(1)) Then Exit Function
                    cfg.CurrentDay = CLng(f(1))
                Case "confirmed"
                    cfg.ConfirmedKey = UnescapeField(f(1))
                Case "header"
                    ReDim Preserve cfg.Headers(0 To cfg.HeaderCount)
                    cfg.Headers(cfg.HeaderCount) = UnescapeField(f(1))
                    cfg.HeaderCount = cfg.HeaderCount + 1
                Case "map"
                    If UBound(f) < 3 Or Not IsCount(f(1)) Then Exit Function
                    If cfg.MapCount >= MAX_FORM_ROWS Then Exit Function
                    sheet = UnescapeField(f(2))
                    If Not ValidSheetName(sheet) Then Exit Function
                    address = NormalizeCellAddress(f(3))
                    If Len(address) = 0 Then Exit Function
                    ReDim Preserve cfg.Maps(0 To cfg.MapCount)
                    cfg.Maps(cfg.MapCount).Column = CLng(f(1))
                    cfg.Maps(cfg.MapCount).SheetName = sheet
                    cfg.Maps(cfg.MapCount).Address = address
                    cfg.MapCount = cfg.MapCount + 1
            End Select
        End If
    Next i
    If Not ValidSheetName(cfg.SourceSheet) Or Len(cfg.SourceAddress) = 0 Then Exit Function
    If cfg.HeaderCount = 0 Or cfg.DateColumn < 0 Or cfg.DateColumn >= cfg.HeaderCount Then Exit Function
    For i = 0 To cfg.MapCount - 1
        If cfg.Maps(i).Column >= cfg.HeaderCount Then Exit Function
    Next i
    ParseConfig = True
End Function
