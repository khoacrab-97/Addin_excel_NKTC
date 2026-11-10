Attribute VB_Name = "modNKTC_PrintCommands"
'==============================================================================
' NHẬT KÝ TC – F5 In nhật ký: hộp thoại frmInNhatKy, in vùng đã chọn cho từng ngày ra máy in hoặc một file PDF.
' In qua workbook tạm (không lưu): mỗi ngày -> hiển thị ngày trên biểu mẫu (cơ chế Chuyển ngày), chép sheet chứa
' vùng in sang workbook tạm và đổi công thức thành giá trị. Xong (kể cả lỗi / Esc): đóng workbook tạm, trả biểu mẫu
' về ngày đang hiển thị trước khi in — file người dùng không bị thêm sheet (CLAUDE.md §4).
'==============================================================================
Option Explicit

' Hằng số Excel (giá trị số).
Private Const XL_WBAT_WORKSHEET As Long = -4167
Private Const XL_TYPE_PDF As Long = 0
Private Const XL_ERROR_HANDLER As Long = 2
Private Const XL_DISABLED As Long = 0
Private Const XL_SHEET_VISIBLE As Long = -1
Private Const XL_PASTE_VALUES As Long = -4163
Private Const ERR_USER_BREAK As Long = 18

' Tên máy in hiện hành của Excel; rỗng nếu máy chưa cài máy in.
Public Function ActivePrinterName() As String
    On Error Resume Next
    ActivePrinterName = Application.ActivePrinter
End Function

'------------------------------------------------------------------------------
' In nhật ký (nút In nhật ký, nhóm In).
'------------------------------------------------------------------------------
Public Sub CmdInNhatKy()
    Dim wb As Workbook, days() As NKTC_SourceDay, dayCount As Long, currentDay As Long, formSheet As Worksheet
    Dim message As String, dayNums() As Long, i As Long, initial As String, f As frmInNhatKy
    Dim accepted As Boolean, printRange As Range, printMode As Long, fromDay As Long, toDay As Long
    Dim toPdf As Boolean, copies As Long, picked() As Long, pickedCount As Long, pdfPath As String

    If ActiveWorkbook Is Nothing Then
        ShowInfo "Hãy mở một file Excel trước.", vbExclamation
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    If Not LoadPrintDays(wb, days, dayCount, currentDay, formSheet, message) Then
        ShowInfo message, vbExclamation
        Exit Sub
    End If
    If dayCount = 0 Then
        ShowInfo "Vùng nguồn không có ngày hợp lệ để in.", vbExclamation
        Exit Sub
    End If
    ReDim dayNums(0 To dayCount - 1)
    For i = 0 To dayCount - 1
        dayNums(i) = days(i).DayNum
    Next i
    initial = DefaultPrintAreaText(formSheet)

    Set f = New frmInNhatKy
    f.Init wb, initial, dayNums, dayCount, currentDay
    Do
        f.Show
        If Len(f.PickRequest) = 0 Then Exit Do
        f.ApplyPick PickRange("In nhật ký", "Chọn vùng sẽ in cho mỗi ngày (biểu mẫu hoặc mẫu in tự dựng, " & _
            "một vùng liên tục trong file này).", f.PickCurrent)
    Loop
    accepted = f.Accepted
    If accepted Then
        Set printRange = f.PrintRange
        printMode = f.Mode
        fromDay = f.FromDay
        toDay = f.ToDay
        toPdf = f.ToPdf
        copies = f.Copies
    End If
    Unload f
    Set f = Nothing
    If Not accepted Then Exit Sub

    message = SelectPrintDays(days, dayCount, printMode, fromDay, toDay, currentDay, picked, pickedCount)
    If Len(message) > 0 Then
        ShowInfo message, vbExclamation
        Exit Sub
    End If
    If pickedCount > PRINT_CONFIRM_DAYS Then
        If ShowMessage("Sẽ in " & pickedCount & " ngày. Tiếp tục?", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If
    If toPdf Then
        pdfPath = AskPdfPath(wb, days(picked(0)).DayNum, days(picked(pickedCount - 1)).DayNum)
        If Len(pdfPath) = 0 Then Exit Sub
    End If

    RunPrint wb, days, picked, pickedCount, printRange, toPdf, copies, pdfPath, currentDay
End Sub

' Vùng in gợi ý: Print Area của sheet biểu mẫu; chưa đặt -> vùng có dữ liệu của sheet đó. Rỗng nếu không có sheet.
Private Function DefaultPrintAreaText(ByVal formSheet As Worksheet) As String
    Dim area As String, r As Range
    If formSheet Is Nothing Then Exit Function
    On Error Resume Next
    area = formSheet.PageSetup.PrintArea
    If Len(area) > 0 Then Set r = formSheet.Range(area)
    On Error GoTo 0
    If r Is Nothing Then Set r = formSheet.UsedRange
    DefaultPrintAreaText = RefText(r)
End Function

' Hỏi nơi lưu file PDF (gợi ý: thư mục của file nhật ký). Rỗng nếu người dùng bấm Cancel.
Private Function AskPdfPath(ByVal wb As Workbook, ByVal firstDay As Long, ByVal lastDay As Long) As String
    Dim initial As String, answer As Variant
    initial = PrintPdfFileName(firstDay, lastDay)
    ' File lưu trên OneDrive/SharePoint có Path dạng https://… -> chỉ gợi ý tên file.
    If Len(wb.Path) > 0 And InStr(1, wb.Path, "://") = 0 Then initial = wb.Path & Application.PathSeparator & initial
    answer = Application.GetSaveAsFilename(InitialFileName:=initial, FileFilter:="PDF (*.pdf), *.pdf")
    If VarType(answer) = vbBoolean Then Exit Function
    AskPdfPath = CStr(answer)
End Function

' Đổi công thức của bản sao (dst, trong workbook tạm) thành giá trị lấy từ sheet gốc (src, đã tính trong file nhật ký):
' bản sao tính lại trong workbook tạm có thể ra kết quả khác (INDIRECT, CELL("filename")…). Dán giá trị giữ nguyên kiểu
' dữ liệu, ô gộp, công thức mảng. Gọi ngay sau khi chép sheet.
Private Sub FreezeValues(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim used As Range, anyFormula As Variant
    Set used = src.UsedRange
    anyFormula = used.HasFormula          ' True / False / Null (chỉ một phần có công thức)
    If Not IsNull(anyFormula) Then
        If Not anyFormula Then Exit Sub
    End If
    used.Copy
    dst.Range(used.Address).PasteSpecial Paste:=XL_PASTE_VALUES
    Application.CutCopyMode = False
End Sub

' In các ngày picked(0..pickedCount-1) (chỉ số trong days()) qua workbook tạm rồi báo kết quả.
' originalDay: ngày đang hiển thị trước khi in (0 nếu chưa có) — được khôi phục cả khi lỗi hoặc bấm Esc.
Private Sub RunPrint(ByVal wb As Workbook, ByRef days() As NKTC_SourceDay, ByRef picked() As Long, _
                     ByVal pickedCount As Long, ByVal printRange As Range, ByVal toPdf As Boolean, _
                     ByVal copies As Long, ByVal pdfPath As String, ByVal originalDay As Long)
    Dim tmp As Workbook, ws As Worksheet, sh As Worksheet, blank As Worksheet, area As String
    Dim k As Long, dayNum As Long, message As String, style As VbMsgBoxStyle, changed As Boolean, shown As Boolean
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldCancel As Long, result As String, kind As VbMsgBoxStyle

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCancel = Application.EnableCancelKey
    kind = vbExclamation
    On Error GoTo Fail
    Set ws = printRange.Worksheet
    area = printRange.Address               ' $A$1:$I$70 (không có tên sheet)
    Application.ScreenUpdating = False
    Application.EnableCancelKey = XL_ERROR_HANDLER      ' Esc -> lỗi 18 -> dừng, dọn dẹp

    Set tmp = Workbooks.Add(XL_WBAT_WORKSHEET)
    tmp.Date1904 = wb.Date1904
    Set blank = tmp.Worksheets(1)
    For k = 0 To pickedCount - 1
        dayNum = days(picked(k)).DayNum
        Application.StatusBar = "Nhật ký TC: chuẩn bị ngày " & (k + 1) & "/" & pickedCount & " (" & FormatDay(dayNum) & ")…"
        changed = True
        ' Không cho Esc ngắt giữa chừng việc ghi biểu mẫu / lưu cấu hình (SaveConfig xóa rồi thêm lại CustomXMLPart).
        Application.EnableCancelKey = XL_DISABLED
        shown = ShowDayOnForm(wb, dayNum, message, style)
        Application.EnableCancelKey = XL_ERROR_HANDLER
        If Not shown Then
            If k = 0 Then changed = False   ' chưa ghi gì lên biểu mẫu
            result = "Đã dừng in: biểu mẫu chưa nhận được dữ liệu ngày " & FormatDay(dayNum) & "."
            If Len(message) > 0 Then result = result & vbLf & vbLf & message
            GoTo Cleanup
        End If
        Application.Calculate               ' mẫu in tự dựng có công thức tham chiếu ô biểu mẫu
        ' Tắt hộp hỏi trùng tên (Name) khi chép sheet nhiều lần vào cùng workbook tạm.
        Application.DisplayAlerts = False
        ws.Copy After:=tmp.Sheets(tmp.Sheets.Count)
        Application.DisplayAlerts = oldAlerts
        Set sh = tmp.Sheets(tmp.Sheets.Count)
        FreezeValues ws, sh
        sh.Visible = XL_SHEET_VISIBLE
        sh.Name = PrintSheetName(dayNum)
        sh.PageSetup.PrintArea = area
    Next k
    Application.DisplayAlerts = False
    blank.Delete
    Application.DisplayAlerts = oldAlerts

    If toPdf Then
        Application.StatusBar = "Nhật ký TC: đang xuất PDF…"
        tmp.ExportAsFixedFormat Type:=XL_TYPE_PDF, Filename:=pdfPath, IgnorePrintAreas:=False, OpenAfterPublish:=False
        result = "Đã xuất " & pickedCount & " ngày ra file:" & vbLf & pdfPath
    Else
        Application.StatusBar = "Nhật ký TC: đang gửi tới máy in…"
        tmp.PrintOut Copies:=copies, Collate:=True
        result = "Đã gửi " & pickedCount & " ngày tới máy in " & ActivePrinterName() & "."
    End If
    kind = vbInformation

Cleanup:
    On Error Resume Next
    Application.EnableCancelKey = XL_DISABLED
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    If Not tmp Is Nothing Then tmp.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    If changed And originalDay > 0 Then
        message = ""
        If Not ShowDayOnForm(wb, originalDay, message, style) Then
            result = result & vbLf & vbLf & "Không trả biểu mẫu về ngày " & FormatDay(originalDay) & " được."
            If Len(message) > 0 Then result = result & " " & message
        End If
    End If
    wb.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel
    On Error GoTo 0
    ShowInfo result, kind
    Exit Sub

Fail:
    Application.EnableCancelKey = XL_DISABLED
    If Err.Number = ERR_USER_BREAK Then
        result = "Đã dừng in theo yêu cầu."
        kind = vbInformation
    Else
        result = "Không in được nhật ký: " & Err.Description
    End If
    Resume Cleanup
End Sub
