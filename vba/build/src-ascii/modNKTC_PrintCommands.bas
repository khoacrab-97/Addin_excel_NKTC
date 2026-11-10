Attribute VB_Name = "modNKTC_PrintCommands"
'==============================================================================
' NHAT KY TC ? F5 In nhat ky: hop thoai frmInNhatKy, in vung da chon cho tung ngay ra may in hoac mot file PDF.
' In qua workbook tam (khong luu): moi ngay -> hien thi ngay tren bieu mau (co che Chuyen ngay), chep sheet chua
' vung in sang workbook tam va doi cong thuc thanh gia tri. Xong (ke ca loi / Esc): dong workbook tam, tra bieu mau
' ve ngay dang hien thi truoc khi in ? file nguoi dung khong bi them sheet (CLAUDE.md ?4).
'==============================================================================
Option Explicit

' Hang so Excel (gia tri so).
Private Const XL_WBAT_WORKSHEET As Long = -4167
Private Const XL_TYPE_PDF As Long = 0
Private Const XL_ERROR_HANDLER As Long = 2
Private Const XL_DISABLED As Long = 0
Private Const XL_SHEET_VISIBLE As Long = -1
Private Const XL_PASTE_VALUES As Long = -4163
Private Const ERR_USER_BREAK As Long = 18

' Ten may in hien hanh cua Excel; rong neu may chua cai may in.
Public Function ActivePrinterName() As String
    On Error Resume Next
    ActivePrinterName = Application.ActivePrinter
End Function

'------------------------------------------------------------------------------
' In nhat ky (nut In nhat ky, nhom In).
'------------------------------------------------------------------------------
Public Sub CmdInNhatKy()
    Dim wb As Workbook, days() As NKTC_SourceDay, dayCount As Long, currentDay As Long, formSheet As Worksheet
    Dim message As String, dayNums() As Long, i As Long, initial As String, f As frmInNhatKy
    Dim accepted As Boolean, printRange As Range, printMode As Long, fromDay As Long, toDay As Long
    Dim toPdf As Boolean, copies As Long, picked() As Long, pickedCount As Long, pdfPath As String

    If ActiveWorkbook Is Nothing Then
        ShowInfo U("H\u00e3y m\u1edf m\u1ed9t file Excel tr\u01b0\u1edbc."), vbExclamation
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    If Not LoadPrintDays(wb, days, dayCount, currentDay, formSheet, message) Then
        ShowInfo message, vbExclamation
        Exit Sub
    End If
    If dayCount = 0 Then
        ShowInfo U("V\u00f9ng ngu\u1ed3n kh\u00f4ng c\u00f3 ng\u00e0y h\u1ee3p l\u1ec7 \u0111\u1ec3 in."), vbExclamation
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
        f.ApplyPick PickRange(U("In nh\u1eadt k\u00fd"), U("Ch\u1ecdn v\u00f9ng s\u1ebd in cho m\u1ed7i ng\u00e0y (bi\u1ec3u m\u1eabu ho\u1eb7c m\u1eabu in t\u1ef1 d\u1ef1ng, ") & _
            U("m\u1ed9t v\u00f9ng li\u00ean t\u1ee5c trong file n\u00e0y)."), f.PickCurrent)
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
        If ShowMessage(U("S\u1ebd in ") & pickedCount & U(" ng\u00e0y. Ti\u1ebfp t\u1ee5c?"), vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If
    If toPdf Then
        pdfPath = AskPdfPath(wb, days(picked(0)).DayNum, days(picked(pickedCount - 1)).DayNum)
        If Len(pdfPath) = 0 Then Exit Sub
    End If

    RunPrint wb, days, picked, pickedCount, printRange, toPdf, copies, pdfPath, currentDay
End Sub

' Vung in goi y: Print Area cua sheet bieu mau; chua dat -> vung co du lieu cua sheet do. Rong neu khong co sheet.
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

' Hoi noi luu file PDF (goi y: thu muc cua file nhat ky). Rong neu nguoi dung bam Cancel.
Private Function AskPdfPath(ByVal wb As Workbook, ByVal firstDay As Long, ByVal lastDay As Long) As String
    Dim initial As String, answer As Variant
    initial = PrintPdfFileName(firstDay, lastDay)
    ' File luu tren OneDrive/SharePoint co Path dang https://? -> chi goi y ten file.
    If Len(wb.Path) > 0 And InStr(1, wb.Path, "://") = 0 Then initial = wb.Path & Application.PathSeparator & initial
    answer = Application.GetSaveAsFilename(InitialFileName:=initial, FileFilter:="PDF (*.pdf), *.pdf")
    If VarType(answer) = vbBoolean Then Exit Function
    AskPdfPath = CStr(answer)
End Function

' Doi cong thuc cua sheet tam thanh gia tri (dan gia tri len chinh no: giu nguyen kieu du lieu, o gop, cong thuc mang).
Private Sub FreezeValues(ByVal sh As Worksheet)
    Dim used As Range, anyFormula As Variant
    Set used = sh.UsedRange
    anyFormula = used.HasFormula          ' True / False / Null (chi mot phan co cong thuc)
    If Not IsNull(anyFormula) Then
        If Not anyFormula Then Exit Sub
    End If
    used.Copy
    used.PasteSpecial Paste:=XL_PASTE_VALUES
    Application.CutCopyMode = False
End Sub

' In cac ngay picked(0..pickedCount-1) (chi so trong days()) qua workbook tam roi bao ket qua.
' originalDay: ngay dang hien thi truoc khi in (0 neu chua co) ? duoc khoi phuc ca khi loi hoac bam Esc.
Private Sub RunPrint(ByVal wb As Workbook, ByRef days() As NKTC_SourceDay, ByRef picked() As Long, _
                     ByVal pickedCount As Long, ByVal printRange As Range, ByVal toPdf As Boolean, _
                     ByVal copies As Long, ByVal pdfPath As String, ByVal originalDay As Long)
    Dim tmp As Workbook, ws As Worksheet, sh As Worksheet, blank As Worksheet, area As String
    Dim k As Long, dayNum As Long, message As String, style As VbMsgBoxStyle, changed As Boolean
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldCancel As Long, result As String, kind As VbMsgBoxStyle

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCancel = Application.EnableCancelKey
    kind = vbExclamation
    On Error GoTo Fail
    Set ws = printRange.Worksheet
    area = printRange.Address               ' $A$1:$I$70 (khong co ten sheet)
    Application.ScreenUpdating = False
    Application.EnableCancelKey = XL_ERROR_HANDLER      ' Esc -> loi 18 -> dung, don dep

    Set tmp = Workbooks.Add(XL_WBAT_WORKSHEET)
    tmp.Date1904 = wb.Date1904
    Set blank = tmp.Worksheets(1)
    For k = 0 To pickedCount - 1
        dayNum = days(picked(k)).DayNum
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: chu\u1ea9n b\u1ecb ng\u00e0y ") & (k + 1) & "/" & pickedCount & " (" & FormatDay(dayNum) & U(")\u2026")
        changed = True
        If Not ShowDayOnForm(wb, dayNum, message, style) Then
            If k = 0 Then changed = False   ' chua ghi gi len bieu mau
            result = U("\u0110\u00e3 d\u1eebng in: bi\u1ec3u m\u1eabu ch\u01b0a nh\u1eadn \u0111\u01b0\u1ee3c d\u1eef li\u1ec7u ng\u00e0y ") & FormatDay(dayNum) & "."
            If Len(message) > 0 Then result = result & vbLf & vbLf & message
            GoTo Cleanup
        End If
        Application.Calculate               ' mau in tu dung co cong thuc tham chieu o bieu mau
        ' Tat hop hoi trung ten (Name) khi chep sheet nhieu lan vao cung workbook tam.
        Application.DisplayAlerts = False
        ws.Copy After:=tmp.Sheets(tmp.Sheets.Count)
        Application.DisplayAlerts = oldAlerts
        Set sh = tmp.Sheets(tmp.Sheets.Count)
        sh.Visible = XL_SHEET_VISIBLE
        sh.Name = PrintSheetName(dayNum)
        FreezeValues sh
        sh.PageSetup.PrintArea = area
    Next k
    Application.DisplayAlerts = False
    blank.Delete
    Application.DisplayAlerts = oldAlerts

    If toPdf Then
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: \u0111ang xu\u1ea5t PDF\u2026")
        tmp.ExportAsFixedFormat Type:=XL_TYPE_PDF, Filename:=pdfPath, IgnorePrintAreas:=False, OpenAfterPublish:=False
        result = U("\u0110\u00e3 xu\u1ea5t ") & pickedCount & U(" ng\u00e0y ra file:") & vbLf & pdfPath
    Else
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: \u0111ang g\u1eedi t\u1edbi m\u00e1y in\u2026")
        tmp.PrintOut Copies:=copies, Collate:=True
        result = U("\u0110\u00e3 g\u1eedi ") & pickedCount & U(" ng\u00e0y t\u1edbi m\u00e1y in ") & ActivePrinterName() & "."
    End If
    kind = vbInformation

Cleanup:
    On Error Resume Next
    Application.EnableCancelKey = XL_DISABLED
    If Not tmp Is Nothing Then tmp.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    If changed And originalDay > 0 Then
        message = ""
        If Not ShowDayOnForm(wb, originalDay, message, style) Then
            result = result & vbLf & vbLf & U("Kh\u00f4ng tr\u1ea3 bi\u1ec3u m\u1eabu v\u1ec1 ng\u00e0y ") & FormatDay(originalDay) & U(" \u0111\u01b0\u1ee3c.")
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
    If Err.Number = ERR_USER_BREAK Then
        result = U("\u0110\u00e3 d\u1eebng in theo y\u00eau c\u1ea7u.")
        kind = vbInformation
    Else
        result = U("Kh\u00f4ng in \u0111\u01b0\u1ee3c nh\u1eadt k\u00fd: ") & Err.Description
    End If
    Resume Cleanup
End Sub
