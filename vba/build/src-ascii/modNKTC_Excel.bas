Attribute VB_Name = "modNKTC_Excel"
'==============================================================================
' NHAT KY TC ? lop Excel: doc vung, tao sheet/bang. Tuong ung src/excel/* cua ban Office.js.
' Khong ghi de du lieu co san (CLAUDE.md ?4): chi ghi vao sheet moi hoac vung da kiem tra la trong.
'==============================================================================
Option Explicit

Private Const XL_TOP As Long = -4160
Private Const XL_CENTER As Long = -4108
Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1
Private Const DATE_FORMAT As String = "dd/mm/yyyy"
Private Const TEXT_FORMAT As String = "@"
Private Const TABLE_STYLE As String = "TableStyleMedium2"
Private Const MAX_LISTED_CELLS As Long = 8

' Dia chi day du de hien trong o chon vung: 'Ten sheet'!$A$1:$D$20.
Public Function RefText(ByVal r As Range) As String
    RefText = "'" & Replace(r.Worksheet.Name, "'", "''") & "'!" & r.Address
End Function

' Doi chu trong o chon vung (go tay hoac tu nut Chon) thanh Range; Nothing neu khong hop le.
Public Function TryRange(ByVal text As String) As Range
    text = Trim$(text)
    If Len(text) = 0 Then Exit Function
    On Error Resume Next
    Set TryRange = Application.Range(text)
    On Error GoTo 0
End Function

' Hop chon vung cua Excel (Application.InputBox Type:=8; khong dung ActiveX/RefEdit vi bi chan mac dinh
' tren Microsoft 365 moi). Tra ve dia chi day du; rong neu nguoi dung bam Cancel.
Public Function PickRange(ByVal title As String, ByVal prompt As String, ByVal current As String, _
                          Optional ByVal singleCell As Boolean = False) As String
    Dim r As Range
    On Error Resume Next
    If TryRange(current) Is Nothing Then
        Set r = Application.InputBox(prompt, title, Type:=8)
    Else
        Set r = Application.InputBox(prompt, title, current, Type:=8)
    End If
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    ' Chon o gop (merge) hay nhieu o -> lay o dau (tren-trai).
    If singleCell Then Set r = r.Cells(1, 1)
    PickRange = RefText(r)
End Function

' Vung dang chon tren sheet (Nothing neu dang chon hinh, bieu do?).
Public Function SelectedRange() As Range
    If TypeName(Selection) = "Range" Then Set SelectedRange = Selection
End Function

' Vung du lieu cho Xuat theo ngay: chon 1 o -> ca vung co du lieu cua sheet (nguoi dung chi can dat con tro
' trong bang); cat theo used range de khong doc ca trieu o khi chon nguyen cot. msg: ly do neu tra ve Nothing.
Public Function ResolveDataRange(ByVal r As Range, ByRef msg As String) As Range
    Dim used As Range, clipped As Range
    msg = ""
    If r Is Nothing Then msg = U("Ch\u01b0a c\u00f3 v\u00f9ng d\u1eef li\u1ec7u."): Exit Function
    If r.Areas.Count > 1 Then msg = U("Vui l\u00f2ng ch\u1ec9 ch\u1ecdn m\u1ed9t v\u00f9ng d\u1eef li\u1ec7u li\u00ean t\u1ee5c."): Exit Function
    Set used = r.Worksheet.UsedRange
    If r.Cells.CountLarge = 1 Then
        Set clipped = used
    Else
        Set clipped = Application.Intersect(r, used)
    End If
    If clipped Is Nothing Then
        msg = U("V\u00f9ng ") & r.Address(False, False) & U(" tr\u00ean sheet """) & r.Worksheet.Name & U(""" kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u.")
        Exit Function
    End If
    If clipped.Rows.Count < 2 Then
        msg = U("V\u00f9ng c\u1ea7n c\u00f3 d\u00f2ng ti\u00eau \u0111\u1ec1 v\u00e0 \u00edt nh\u1ea5t m\u1ed9t d\u00f2ng c\u00f4ng vi\u1ec7c.")
        Exit Function
    End If
    Set ResolveDataRange = clipped
End Function

' Gia tri vung luon la mang 2 chieu 1-based (Range.Value2 cua 1 o tra ve gia tri don).
Public Function RangeGrid(ByVal r As Range) As Variant
    Dim v As Variant, one(1 To 1, 1 To 1) As Variant
    v = r.Value2
    If IsArray(v) Then
        RangeGrid = v
    Else
        one(1, 1) = v
        RangeGrid = one
    End If
End Function

' Dong tieu de cua vung: mang 1 chieu 0-based.
Public Function HeaderRow(ByVal r As Range) As Variant
    Dim grid As Variant, c As Long, out() As Variant
    grid = RangeGrid(r.Rows(1))
    ReDim out(0 To UBound(grid, 2) - 1)
    For c = 1 To UBound(grid, 2)
        out(c - 1) = grid(1, c)
    Next c
    HeaderRow = out
End Function

' Vung nguon tu mo rong (R2.17): tu o dau (tren-trai) cua r toi dong/cot cuoi co du lieu cua sheet ? them cot
' (ke ca chen giua) hay them ngay o duoi deu duoc tinh, khong phai chon lai vung.
Public Function ExpandSource(ByVal r As Range) As Range
    Dim anchor As Range, used As Range, lastRow As Long, lastCol As Long
    Set anchor = r.Cells(1, 1)
    Set used = anchor.Worksheet.UsedRange
    lastRow = used.Row + used.Rows.Count - 1
    lastCol = used.Column + used.Columns.Count - 1
    If lastRow < anchor.Row Or lastCol < anchor.Column Then
        Set ExpandSource = r
    Else
        Set ExpandSource = anchor.Worksheet.Range(anchor, anchor.Worksheet.Cells(lastRow, lastCol))
    End If
End Function

' Chu cua dong tieu de (0-based): chuoi giu nguyen, so/logic thanh chu, con lai rong.
Public Function HeaderTexts(ByVal r As Range) As String()
    Dim h As Variant, out() As String, i As Long
    h = HeaderRow(r)
    ReDim out(0 To UBound(h))
    For i = 0 To UBound(h)
        out(i) = CellText(h(i))
    Next i
    HeaderTexts = out
End Function

' Sheet co ten (khong phan biet hoa/thuong); Nothing neu khong co.
Public Function FindSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = LCase$(name) Then Set FindSheet = ws: Exit Function
    Next ws
End Function

Public Function SheetNames(ByVal wb As Workbook) As Collection
    Dim names As New Collection, sh As Object
    For Each sh In wb.Sheets
        names.Add sh.Name
    Next sh
    Set SheetNames = names
End Function

' Them sheet o cuoi workbook voi ten chua trung (base, base_2, ?).
Private Function AddSheet(ByVal wb As Workbook, ByVal base As String) As Worksheet
    Dim ws As Worksheet, name As String
    name = UniqueSheetName(SheetNames(wb), base)
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    Set AddSheet = ws
End Function

'------------------------------------------------------------------------------
' F1: tao sheet NK_TheoNgay va ghi nhat ky theo ngay (R1.8, R1.9). Tra ve ten sheet.
'------------------------------------------------------------------------------
Public Function WriteDailyLogSheet(ByVal wb As Workbook, ByRef entries() As NKTC_DailyEntry, _
                                   ByVal entryCount As Long, ByVal date1904 As Boolean) As String
    Dim ws As Worksheet, dates() As Variant, texts() As Variant, i As Long
    Set ws = AddSheet(wb, DAILY_SHEET_BASE_NAME)
    With ws.Range("A1:B1")
        .Value2 = Array(U("Ng\u00e0y"), U("N\u1ed9i dung"))
        .Font.Bold = True
    End With
    If entryCount > 0 Then
        ReDim dates(1 To entryCount, 1 To 1)
        ReDim texts(1 To entryCount, 1 To 1)
        For i = 0 To entryCount - 1
            dates(i + 1, 1) = ToWorkbookSerial(entries(i).DayNum, date1904)
            texts(i + 1, 1) = entries(i).Content
        Next i
        ' Dinh dang truoc khi ghi: noi dung bat dau bang "-", "=", "+" khong bi hieu la cong thuc.
        With ws.Range("A2").Resize(entryCount, 1)
            .NumberFormat = DATE_FORMAT
            .Value2 = dates
        End With
        With ws.Range("B2").Resize(entryCount, 1)
            .NumberFormat = TEXT_FORMAT
            .Value2 = texts
            .WrapText = True
        End With
    End If
    ws.Range("A1").Resize(entryCount + 1, 2).VerticalAlignment = XL_TOP
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 85
    ws.Activate
    ws.Range("A1").Select
    WriteDailyLogSheet = ws.Name
End Function

'------------------------------------------------------------------------------
' F4: tao bang nhap tien do STT | Noi dung cong viec | Bat dau | Ket thuc (Excel Table).
' Sheet hien tai: chi tao khi ca vung bang con trong (R4.4). Tra ve chuoi loi; rong neu thanh cong.
'------------------------------------------------------------------------------
Public Function CreateInputTable(ByVal onNewSheet As Boolean, ByVal startCell As Range, ByVal rows As Long, _
                                 ByRef createdAt As String) As String
    Dim ws As Worksheet, area As Range, body As Range, numbers() As Variant, i As Long, lo As Object
    Dim widths As Variant
    If onNewSheet Then
        Set ws = AddSheet(ActiveWorkbook, INPUT_SHEET_BASE_NAME)
        Set area = ws.Range("A1").Resize(rows + 1, 4)
    Else
        Set ws = startCell.Worksheet
        If startCell.Row + rows > ws.Rows.Count Or startCell.Column + 3 > ws.Columns.Count Then
            CreateInputTable = U("B\u1ea3ng v\u01b0\u1ee3t qu\u00e1 gi\u1edbi h\u1ea1n sheet c\u1ee7a Excel. Ch\u1ecdn \u00f4 b\u1eaft \u0111\u1ea7u kh\u00e1c ho\u1eb7c gi\u1ea3m s\u1ed1 d\u00f2ng.")
            Exit Function
        End If
        Set area = startCell.Cells(1, 1).Resize(rows + 1, 4)
        CreateInputTable = AreaProblem(ws, area)
        If Len(CreateInputTable) > 0 Then Exit Function
    End If

    Set body = area.Offset(1, 0).Resize(rows, 4)
    body.Columns(2).NumberFormat = TEXT_FORMAT
    body.Columns(3).NumberFormat = DATE_FORMAT
    body.Columns(4).NumberFormat = DATE_FORMAT
    body.Columns(2).WrapText = True
    area.Rows(1).Value2 = InputTableHeaders()
    ReDim numbers(1 To rows, 1 To 1)
    For i = 1 To rows
        numbers(i, 1) = i
    Next i
    body.Columns(1).Value2 = numbers

    Set lo = ws.ListObjects.Add(XL_SRC_RANGE, area, , XL_YES)
    lo.TableStyle = TABLE_STYLE
    area.VerticalAlignment = XL_TOP
    body.Columns(1).HorizontalAlignment = XL_CENTER
    widths = Array(5, 60, 12, 12)
    For i = 0 To 3
        ws.Columns(area.Column + i).ColumnWidth = widths(i)
    Next i
    ws.Activate
    body.Cells(1, 2).Select
    createdAt = ws.Name & "!" & area.Address(False, False)
End Function

' Ly do khong tao duoc bang tai vung tren sheet hien tai; rong neu duoc.
Private Function AreaProblem(ByVal ws As Worksheet, ByVal area As Range) As String
    Dim c As Range, listed As String, n As Long, lo As Object, merged As Variant
    If ws.ProtectContents Then
        AreaProblem = "Sheet """ & ws.Name & U(""" \u0111ang \u0111\u01b0\u1ee3c b\u1ea3o v\u1ec7 (Protect Sheet). B\u1ecf b\u1ea3o v\u1ec7 ho\u1eb7c t\u1ea1o tr\u00ean sheet m\u1edbi.")
        Exit Function
    End If
    For Each c In area.Cells
        If c.HasFormula Or Not IsEmpty(c.Value2) Then
            n = n + 1
            If n <= MAX_LISTED_CELLS Then
                If Len(listed) > 0 Then listed = listed & ", "
                listed = listed & c.Address(False, False)
            End If
        End If
    Next c
    If n > 0 Then
        If n > MAX_LISTED_CELLS Then listed = listed & U(" v\u00e0 ") & (n - MAX_LISTED_CELLS) & U(" \u00f4 kh\u00e1c")
        AreaProblem = U("V\u00f9ng ") & area.Address(False, False) & U(" \u0111ang c\u00f3 d\u1eef li\u1ec7u (") & listed & _
            U("). Add-in kh\u00f4ng ghi \u0111\u00e8 \u2014 ch\u1ecdn \u00f4 b\u1eaft \u0111\u1ea7u kh\u00e1c ho\u1eb7c t\u1ea1o tr\u00ean sheet m\u1edbi.")
        Exit Function
    End If
    ' MergeCells: Null khi vung co ca o gop va o thuong.
    merged = area.MergeCells
    If IsNull(merged) Then merged = True
    If merged Then
        AreaProblem = U("V\u00f9ng ") & area.Address(False, False) & U(" c\u00f3 \u00f4 g\u1ed9p (merge). Ch\u1ecdn \u00f4 b\u1eaft \u0111\u1ea7u kh\u00e1c ho\u1eb7c t\u1ea1o tr\u00ean sheet m\u1edbi.")
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If Not Application.Intersect(lo.Range, area) Is Nothing Then
            AreaProblem = U("V\u00f9ng ") & area.Address(False, False) & U(" ch\u1ed3ng l\u00ean b\u1ea3ng """) & lo.Name & U(""" c\u00f3 s\u1eb5n.")
            Exit Function
        End If
    Next lo
End Function
