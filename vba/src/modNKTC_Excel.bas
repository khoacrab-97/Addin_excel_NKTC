Attribute VB_Name = "modNKTC_Excel"
'==============================================================================
' NHẬT KÝ TC – lớp Excel: đọc vùng, tạo sheet/bảng. Tương ứng src/excel/* của bản Office.js.
' Không ghi đè dữ liệu có sẵn (CLAUDE.md §4): chỉ ghi vào sheet mới hoặc vùng đã kiểm tra là trống.
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

' Địa chỉ đầy đủ để hiện trong ô chọn vùng: 'Tên sheet'!$A$1:$D$20.
Public Function RefText(ByVal r As Range) As String
    RefText = "'" & Replace(r.Worksheet.Name, "'", "''") & "'!" & r.Address
End Function

' Đổi chữ trong ô chọn vùng (RefEdit hoặc gõ tay) thành Range; Nothing nếu không hợp lệ.
Public Function TryRange(ByVal text As String) As Range
    text = Trim$(text)
    If Len(text) = 0 Then Exit Function
    On Error Resume Next
    Set TryRange = Application.Range(text)
    On Error GoTo 0
End Function

' Vùng đang chọn trên sheet (Nothing nếu đang chọn hình, biểu đồ…).
Public Function SelectedRange() As Range
    If TypeName(Selection) = "Range" Then Set SelectedRange = Selection
End Function

' Vùng dữ liệu cho Xuất theo ngày: chọn 1 ô -> cả vùng có dữ liệu của sheet (người dùng chỉ cần đặt con trỏ
' trong bảng); cắt theo used range để không đọc cả triệu ô khi chọn nguyên cột. msg: lý do nếu trả về Nothing.
Public Function ResolveDataRange(ByVal r As Range, ByRef msg As String) As Range
    Dim used As Range, clipped As Range
    msg = ""
    If r Is Nothing Then msg = "Chưa có vùng dữ liệu.": Exit Function
    If r.Areas.Count > 1 Then msg = "Vui lòng chỉ chọn một vùng dữ liệu liên tục.": Exit Function
    Set used = r.Worksheet.UsedRange
    If r.Cells.CountLarge = 1 Then
        Set clipped = used
    Else
        Set clipped = Application.Intersect(r, used)
    End If
    If clipped Is Nothing Then
        msg = "Vùng " & r.Address(False, False) & " trên sheet """ & r.Worksheet.Name & """ không có dữ liệu."
        Exit Function
    End If
    If clipped.Rows.Count < 2 Then
        msg = "Vùng cần có dòng tiêu đề và ít nhất một dòng công việc."
        Exit Function
    End If
    Set ResolveDataRange = clipped
End Function

' Giá trị vùng luôn là mảng 2 chiều 1-based (Range.Value2 của 1 ô trả về giá trị đơn).
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

' Dòng tiêu đề của vùng: mảng 1 chiều 0-based.
Public Function HeaderRow(ByVal r As Range) As Variant
    Dim grid As Variant, c As Long, out() As Variant
    grid = RangeGrid(r.Rows(1))
    ReDim out(0 To UBound(grid, 2) - 1)
    For c = 1 To UBound(grid, 2)
        out(c - 1) = grid(1, c)
    Next c
    HeaderRow = out
End Function

Public Function SheetNames(ByVal wb As Workbook) As Collection
    Dim names As New Collection, sh As Object
    For Each sh In wb.Sheets
        names.Add sh.Name
    Next sh
    Set SheetNames = names
End Function

' Thêm sheet ở cuối workbook với tên chưa trùng (base, base_2, …).
Private Function AddSheet(ByVal wb As Workbook, ByVal base As String) As Worksheet
    Dim ws As Worksheet, name As String
    name = UniqueSheetName(SheetNames(wb), base)
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    Set AddSheet = ws
End Function

'------------------------------------------------------------------------------
' F1: tạo sheet NK_TheoNgay và ghi nhật ký theo ngày (R1.8, R1.9). Trả về tên sheet.
'------------------------------------------------------------------------------
Public Function WriteDailyLogSheet(ByVal wb As Workbook, ByRef entries() As NKTC_DailyEntry, _
                                   ByVal entryCount As Long, ByVal date1904 As Boolean) As String
    Dim ws As Worksheet, dates() As Variant, texts() As Variant, i As Long
    Set ws = AddSheet(wb, DAILY_SHEET_BASE_NAME)
    With ws.Range("A1:B1")
        .Value2 = Array("Ngày", "Nội dung")
        .Font.Bold = True
    End With
    If entryCount > 0 Then
        ReDim dates(1 To entryCount, 1 To 1)
        ReDim texts(1 To entryCount, 1 To 1)
        For i = 0 To entryCount - 1
            dates(i + 1, 1) = ToWorkbookSerial(entries(i).DayNum, date1904)
            texts(i + 1, 1) = entries(i).Content
        Next i
        ' Định dạng trước khi ghi: nội dung bắt đầu bằng "-", "=", "+" không bị hiểu là công thức.
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
' F4: tạo bảng nhập tiến độ STT | Nội dung công việc | Bắt đầu | Kết thúc (Excel Table).
' Sheet hiện tại: chỉ tạo khi cả vùng bảng còn trống (R4.4). Trả về chuỗi lỗi; rỗng nếu thành công.
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
            CreateInputTable = "Bảng vượt quá giới hạn sheet của Excel. Chọn ô bắt đầu khác hoặc giảm số dòng."
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

' Lý do không tạo được bảng tại vùng trên sheet hiện tại; rỗng nếu được.
Private Function AreaProblem(ByVal ws As Worksheet, ByVal area As Range) As String
    Dim c As Range, listed As String, n As Long, lo As Object, merged As Variant
    If ws.ProtectContents Then
        AreaProblem = "Sheet """ & ws.Name & """ đang được bảo vệ (Protect Sheet). Bỏ bảo vệ hoặc tạo trên sheet mới."
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
        If n > MAX_LISTED_CELLS Then listed = listed & " và " & (n - MAX_LISTED_CELLS) & " ô khác"
        AreaProblem = "Vùng " & area.Address(False, False) & " đang có dữ liệu (" & listed & _
            "). Add-in không ghi đè — chọn ô bắt đầu khác hoặc tạo trên sheet mới."
        Exit Function
    End If
    ' MergeCells: Null khi vùng có cả ô gộp và ô thường.
    merged = area.MergeCells
    If IsNull(merged) Then merged = True
    If merged Then
        AreaProblem = "Vùng " & area.Address(False, False) & " có ô gộp (merge). Chọn ô bắt đầu khác hoặc tạo trên sheet mới."
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If Not Application.Intersect(lo.Range, area) Is Nothing Then
            AreaProblem = "Vùng " & area.Address(False, False) & " chồng lên bảng """ & lo.Name & """ có sẵn."
            Exit Function
        End If
    Next lo
End Function
