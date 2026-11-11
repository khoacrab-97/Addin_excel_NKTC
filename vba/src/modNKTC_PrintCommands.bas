Attribute VB_Name = "modNKTC_PrintCommands"
'==============================================================================
' NHẬT KÝ TC – F5 In nhật ký: hộp thoại frmInNhatKy, in vùng đã chọn cho từng ngày ra máy in hoặc một file PDF.
' In qua workbook tạm (không lưu): mỗi ngày -> hiển thị ngày trên biểu mẫu (cơ chế Chuyển ngày), chép sheet chứa
' vùng in sang workbook tạm và đổi công thức thành giá trị. Xong (kể cả lỗi / Esc): đóng workbook tạm, trả biểu mẫu
' về ngày đang hiển thị trước khi in — file người dùng không bị thêm sheet (CLAUDE.md §4).
' Chỉ in đúng ngày đang hiển thị: in biểu mẫu như đang thấy (giữ nội dung sửa tay), không ghi lại từ nguồn.
'==============================================================================
Option Explicit

' Hằng số Excel (giá trị số).
Private Const XL_TYPE_PDF As Long = 0
Private Const XL_ERROR_HANDLER As Long = 2
Private Const XL_DISABLED As Long = 0
Private Const XL_SHEET_VISIBLE As Long = -1
Private Const XL_CELL_TYPE_FORMULAS As Long = -4123
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
    Dim cfg As NKTC_Config, sourceSheet As String

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
    If wb.ProtectStructure Then
        ShowInfo "Workbook đang khóa cấu trúc (Protect Workbook) — không chép được sheet để in. Bỏ khóa rồi in lại.", _
            vbExclamation
        Exit Sub
    End If
    If LoadConfig(wb, cfg) Then sourceSheet = cfg.SourceSheet     ' sheet nguồn nhật ký: không cho chọn làm vùng in
    ' Biểu mẫu đặt chung sheet với vùng nguồn: không chặn (nếu chặn thì không in được biểu mẫu).
    If StrComp(sourceSheet, formSheet.Name, vbTextCompare) = 0 Then sourceSheet = ""
    ReDim dayNums(0 To dayCount - 1)
    For i = 0 To dayCount - 1
        dayNums(i) = days(i).DayNum
    Next i
    initial = DefaultPrintAreaText(formSheet)

    Set f = New frmInNhatKy
    f.Init wb, initial, dayNums, dayCount, currentDay, sourceSheet
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

' Vùng in gợi ý: Print Area của sheet biểu mẫu; chưa đặt hoặc gồm nhiều vùng -> vùng có dữ liệu của sheet đó.
' Rỗng nếu không có sheet.
Private Function DefaultPrintAreaText(ByVal formSheet As Worksheet) As String
    Dim area As String, r As Range
    If formSheet Is Nothing Then Exit Function
    On Error Resume Next
    area = formSheet.PageSetup.PrintArea
    If Len(area) > 0 Then Set r = formSheet.Range(area)
    On Error GoTo 0
    If Not r Is Nothing Then
        If r.Areas.Count > 1 Then Set r = Nothing
    End If
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
' bản sao tính lại trong workbook tạm có thể ra kết quả khác (INDIRECT, CELL("filename")…). Chỉ ghi vào ô có công thức
' (không dán cả vùng: Excel từ chối dán đè vùng có ô gộp), không dùng Clipboard, không đụng các ô khác. Gọi ngay sau
' khi chép sheet.
Private Sub FreezeValues(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim formulas As Range, c As Range, block As Range, probe As Object, spills As Boolean
    On Error Resume Next
    Set formulas = src.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS)   ' lỗi khi không có ô công thức nào
    On Error GoTo 0
    If formulas Is Nothing Then Exit Sub
    For Each c In formulas.Cells
        Set block = Nothing
        If c.HasArray Then
            ' Công thức mảng Ctrl+Shift+Enter: ghi cả mảng một lần, tại ô đầu của mảng.
            If c.Address = c.CurrentArray.Cells(1, 1).Address Then Set block = c.CurrentArray
        Else
            ' Công thức mảng động (spill, Excel 365): ghi cả vùng tràn. Excel cũ không có HasSpill -> chỉ ghi ô đó.
            Set probe = c
            spills = False
            On Error Resume Next
            spills = probe.HasSpill
            If spills Then Set block = probe.SpillingToRange
            On Error GoTo 0
            If block Is Nothing Then Set block = c
        End If
        If Not block Is Nothing Then WriteFrozen dst.Range(block.Address), block.Value2
    Next c
End Sub

' Ghi giá trị đã tính vào bản sao. Chữ được thêm dấu ' ở đầu để Excel không đọc lại ("- ...", "= ..." thành công thức,
' "001" thành số, "05/06" thành ngày).
Private Sub WriteFrozen(ByVal target As Range, ByVal values As Variant)
    Dim i As Long, j As Long
    If IsArray(values) Then
        For i = LBound(values, 1) To UBound(values, 1)
            For j = LBound(values, 2) To UBound(values, 2)
                If VarType(values(i, j)) = vbString Then values(i, j) = "'" & values(i, j)
            Next j
        Next i
    ElseIf VarType(values) = vbString Then
        values = "'" & values
    End If
    If target.Cells(1, 1).HasArray Then target.ClearContents     ' mảng Ctrl+Shift+Enter: xóa cả mảng rồi mới ghi được
    target.Value2 = values
End Sub

' In các ngày picked(0..pickedCount-1) (chỉ số trong days()) qua workbook tạm rồi báo kết quả.
' originalDay: ngày đang hiển thị trước khi in (0 nếu chưa có) — được khôi phục cả khi lỗi hoặc bấm Esc.
' Chỉ in một ngày và đó là ngày đang hiển thị -> in biểu mẫu như đang thấy (không ghi lại, không khôi phục).
Private Sub RunPrint(ByVal wb As Workbook, ByRef days() As NKTC_SourceDay, ByRef picked() As Long, _
                     ByVal pickedCount As Long, ByVal printRange As Range, ByVal toPdf As Boolean, _
                     ByVal copies As Long, ByVal pdfPath As String, ByVal originalDay As Long)
    Dim tmp As Workbook, newBook As Workbook, ws As Worksheet, sh As Worksheet, area As String
    Dim k As Long, dayNum As Long, message As String, style As VbMsgBoxStyle, changed As Boolean, shown As Boolean
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldCancel As Long, result As String, kind As VbMsgBoxStyle
    Dim keepDisplayed As Boolean, lastShown As Long

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCancel = Application.EnableCancelKey
    kind = vbExclamation
    On Error GoTo Fail
    Set ws = printRange.Worksheet
    area = printRange.Address               ' $A$1:$I$70 (không có tên sheet)
    If pickedCount = 1 And originalDay > 0 Then keepDisplayed = (days(picked(0)).DayNum = originalDay)
    Application.ScreenUpdating = False
    Application.EnableCancelKey = XL_ERROR_HANDLER      ' Esc -> lỗi 18 -> dừng, dọn dẹp

    For k = 0 To pickedCount - 1
        dayNum = days(picked(k)).DayNum
        Application.StatusBar = "Nhật ký TC: chuẩn bị ngày " & (k + 1) & "/" & pickedCount & " (" & FormatDay(dayNum) & ")…"
        If Not keepDisplayed Then
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
            lastShown = dayNum
        End If
        Application.Calculate               ' mẫu in tự dựng có công thức tham chiếu ô biểu mẫu
        ' Tắt hộp hỏi trùng tên (Name) khi chép sheet nhiều lần vào cùng workbook tạm.
        Application.DisplayAlerts = False
        If tmp Is Nothing Then
            ' Ngày đầu: chép sheet ra workbook mới -> workbook tạm giữ theme / style của file nhật ký.
            ' Không cho Esc ngắt giữa lúc workbook mới đã có mà chưa gán vào tmp (sẽ bị bỏ lại không đóng).
            Application.EnableCancelKey = XL_DISABLED
            ws.Copy
            Application.DisplayAlerts = oldAlerts
            Set newBook = ActiveWorkbook
            If newBook Is Nothing Then
                result = "Không tạo được workbook tạm để in."
                GoTo Cleanup
            End If
            If newBook Is wb Then
                result = "Không tạo được workbook tạm để in."
                GoTo Cleanup
            End If
            Set tmp = newBook
            Application.EnableCancelKey = XL_ERROR_HANDLER
            tmp.Date1904 = wb.Date1904
        Else
            ws.Copy After:=tmp.Sheets(tmp.Sheets.Count)
            Application.DisplayAlerts = oldAlerts
        End If
        Set sh = tmp.Sheets(tmp.Sheets.Count)
        ' Chỉ bỏ khóa bản sao trong workbook tạm (sheet của người dùng giữ nguyên). Password:="" -> sheet có mật khẩu
        ' báo lỗi thay vì hiện hộp hỏi mật khẩu.
        If sh.ProtectContents Then
            On Error Resume Next
            sh.Unprotect Password:=""
            On Error GoTo Fail
            If sh.ProtectContents Then
                result = "Sheet chứa vùng in đang khóa (Protect Sheet) có mật khẩu — bỏ khóa rồi in lại."
                GoTo Cleanup
            End If
        End If
        FreezeValues ws, sh
        sh.Visible = XL_SHEET_VISIBLE
        sh.Name = PrintSheetName(dayNum)
        sh.PageSetup.PrintArea = area
    Next k

    If toPdf Then
        Application.StatusBar = "Nhật ký TC: đang xuất PDF…"
        tmp.ExportAsFixedFormat Type:=XL_TYPE_PDF, Filename:=pdfPath, IgnorePrintAreas:=False, OpenAfterPublish:=False
        result = "Đã xuất " & pickedCount & " ngày ra file:" & vbLf & pdfPath
    Else
        Application.StatusBar = "Nhật ký TC: đang gửi tới máy in…"
        tmp.PrintOut Copies:=copies, Collate:=True
        result = "Đã gửi " & pickedCount & " ngày tới máy in " & ActivePrinterName() & "."
    End If
    If originalDay = 0 And lastShown > 0 Then
        result = result & vbLf & "Biểu mẫu đang hiển thị ngày " & FormatDay(lastShown) & "."
    End If
    kind = vbInformation

Cleanup:
    On Error Resume Next
    Application.EnableCancelKey = XL_DISABLED
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
