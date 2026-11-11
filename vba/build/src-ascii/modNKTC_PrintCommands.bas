Attribute VB_Name = "modNKTC_PrintCommands"
'==============================================================================
' NHAT KY TC ? F5 In nhat ky: hop thoai frmInNhatKy, in vung da chon cho tung ngay ra may in hoac mot file PDF.
' In qua workbook tam (khong luu): moi ngay -> hien thi ngay tren bieu mau (co che Chuyen ngay), chep sheet chua
' vung in sang workbook tam va doi cong thuc thanh gia tri. Xong (ke ca loi / Esc): dong workbook tam, tra bieu mau
' ve ngay dang hien thi truoc khi in ? file nguoi dung khong bi them sheet (CLAUDE.md ?4).
' Chi in dung ngay dang hien thi: in bieu mau nhu dang thay (giu noi dung sua tay), khong ghi lai tu nguon.
'==============================================================================
Option Explicit

' Hang so Excel (gia tri so).
Private Const XL_TYPE_PDF As Long = 0
Private Const XL_ERROR_HANDLER As Long = 2
Private Const XL_DISABLED As Long = 0
Private Const XL_SHEET_VISIBLE As Long = -1
Private Const XL_CELL_TYPE_FORMULAS As Long = -4123
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
    Dim cfg As NKTC_Config, sourceSheet As String

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
    If wb.ProtectStructure Then
        ShowInfo U("Workbook \u0111ang kh\u00f3a c\u1ea5u tr\u00fac (Protect Workbook) \u2014 kh\u00f4ng ch\u00e9p \u0111\u01b0\u1ee3c sheet \u0111\u1ec3 in. B\u1ecf kh\u00f3a r\u1ed3i in l\u1ea1i."), _
            vbExclamation
        Exit Sub
    End If
    If LoadConfig(wb, cfg) Then sourceSheet = cfg.SourceSheet     ' sheet nguon nhat ky: khong cho chon lam vung in
    ' Bieu mau dat chung sheet voi vung nguon: khong chan (neu chan thi khong in duoc bieu mau).
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

' Vung in goi y: Print Area cua sheet bieu mau; chua dat hoac gom nhieu vung -> vung co du lieu cua sheet do.
' Rong neu khong co sheet.
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

' Doi cong thuc cua ban sao (dst, trong workbook tam) thanh gia tri lay tu sheet goc (src, da tinh trong file nhat ky):
' ban sao tinh lai trong workbook tam co the ra ket qua khac (INDIRECT, CELL("filename")?). Chi ghi vao o co cong thuc
' (khong dan ca vung: Excel tu choi dan de vung co o gop), khong dung Clipboard, khong dung cac o khac. Goi ngay sau
' khi chep sheet.
Private Sub FreezeValues(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim formulas As Range, c As Range, block As Range, probe As Object, spills As Boolean
    On Error Resume Next
    Set formulas = src.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS)   ' loi khi khong co o cong thuc nao
    On Error GoTo 0
    If formulas Is Nothing Then Exit Sub
    For Each c In formulas.Cells
        Set block = Nothing
        If c.HasArray Then
            ' Cong thuc mang Ctrl+Shift+Enter: ghi ca mang mot lan, tai o dau cua mang.
            If c.Address = c.CurrentArray.Cells(1, 1).Address Then Set block = c.CurrentArray
        Else
            ' Cong thuc mang dong (spill, Excel 365): ghi ca vung tran. Excel cu khong co HasSpill -> chi ghi o do.
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

' Ghi gia tri da tinh vao ban sao. Chu duoc them dau ' o dau de Excel khong doc lai ("- ...", "= ..." thanh cong thuc,
' "001" thanh so, "05/06" thanh ngay).
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
    If target.Cells(1, 1).HasArray Then target.ClearContents     ' mang Ctrl+Shift+Enter: xoa ca mang roi moi ghi duoc
    target.Value2 = values
End Sub

' In cac ngay picked(0..pickedCount-1) (chi so trong days()) qua workbook tam roi bao ket qua.
' originalDay: ngay dang hien thi truoc khi in (0 neu chua co) ? duoc khoi phuc ca khi loi hoac bam Esc.
' Chi in mot ngay va do la ngay dang hien thi -> in bieu mau nhu dang thay (khong ghi lai, khong khoi phuc).
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
    area = printRange.Address               ' $A$1:$I$70 (khong co ten sheet)
    If pickedCount = 1 And originalDay > 0 Then keepDisplayed = (days(picked(0)).DayNum = originalDay)
    Application.ScreenUpdating = False
    Application.EnableCancelKey = XL_ERROR_HANDLER      ' Esc -> loi 18 -> dung, don dep

    For k = 0 To pickedCount - 1
        dayNum = days(picked(k)).DayNum
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: chu\u1ea9n b\u1ecb ng\u00e0y ") & (k + 1) & "/" & pickedCount & " (" & FormatDay(dayNum) & U(")\u2026")
        If Not keepDisplayed Then
            changed = True
            ' Khong cho Esc ngat giua chung viec ghi bieu mau / luu cau hinh (SaveConfig xoa roi them lai CustomXMLPart).
            Application.EnableCancelKey = XL_DISABLED
            shown = ShowDayOnForm(wb, dayNum, message, style)
            Application.EnableCancelKey = XL_ERROR_HANDLER
            If Not shown Then
                If k = 0 Then changed = False   ' chua ghi gi len bieu mau
                result = U("\u0110\u00e3 d\u1eebng in: bi\u1ec3u m\u1eabu ch\u01b0a nh\u1eadn \u0111\u01b0\u1ee3c d\u1eef li\u1ec7u ng\u00e0y ") & FormatDay(dayNum) & "."
                If Len(message) > 0 Then result = result & vbLf & vbLf & message
                GoTo Cleanup
            End If
            lastShown = dayNum
        End If
        Application.Calculate               ' mau in tu dung co cong thuc tham chieu o bieu mau
        ' Tat hop hoi trung ten (Name) khi chep sheet nhieu lan vao cung workbook tam.
        Application.DisplayAlerts = False
        If tmp Is Nothing Then
            ' Ngay dau: chep sheet ra workbook moi -> workbook tam giu theme / style cua file nhat ky.
            ' Khong cho Esc ngat giua luc workbook moi da co ma chua gan vao tmp (se bi bo lai khong dong).
            Application.EnableCancelKey = XL_DISABLED
            ws.Copy
            Application.DisplayAlerts = oldAlerts
            Set newBook = ActiveWorkbook
            If newBook Is Nothing Then
                result = U("Kh\u00f4ng t\u1ea1o \u0111\u01b0\u1ee3c workbook t\u1ea1m \u0111\u1ec3 in.")
                GoTo Cleanup
            End If
            If newBook Is wb Then
                result = U("Kh\u00f4ng t\u1ea1o \u0111\u01b0\u1ee3c workbook t\u1ea1m \u0111\u1ec3 in.")
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
        ' Chi bo khoa ban sao trong workbook tam (sheet cua nguoi dung giu nguyen). Password:="" -> sheet co mat khau
        ' bao loi thay vi hien hop hoi mat khau.
        If sh.ProtectContents Then
            On Error Resume Next
            sh.Unprotect Password:=""
            On Error GoTo Fail
            If sh.ProtectContents Then
                result = U("Sheet ch\u1ee9a v\u00f9ng in \u0111ang kh\u00f3a (Protect Sheet) c\u00f3 m\u1eadt kh\u1ea9u \u2014 b\u1ecf kh\u00f3a r\u1ed3i in l\u1ea1i.")
                GoTo Cleanup
            End If
        End If
        FreezeValues ws, sh
        sh.Visible = XL_SHEET_VISIBLE
        sh.Name = PrintSheetName(dayNum)
        sh.PageSetup.PrintArea = area
    Next k

    If toPdf Then
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: \u0111ang xu\u1ea5t PDF\u2026")
        tmp.ExportAsFixedFormat Type:=XL_TYPE_PDF, Filename:=pdfPath, IgnorePrintAreas:=False, OpenAfterPublish:=False
        result = U("\u0110\u00e3 xu\u1ea5t ") & pickedCount & U(" ng\u00e0y ra file:") & vbLf & pdfPath
    Else
        Application.StatusBar = U("Nh\u1eadt k\u00fd TC: \u0111ang g\u1eedi t\u1edbi m\u00e1y in\u2026")
        tmp.PrintOut Copies:=copies, Collate:=True
        result = U("\u0110\u00e3 g\u1eedi ") & pickedCount & U(" ng\u00e0y t\u1edbi m\u00e1y in ") & ActivePrinterName() & "."
    End If
    If originalDay = 0 And lastShown > 0 Then
        result = result & vbLf & U("Bi\u1ec3u m\u1eabu \u0111ang hi\u1ec3n th\u1ecb ng\u00e0y ") & FormatDay(lastShown) & "."
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
    Application.EnableCancelKey = XL_DISABLED
    If Err.Number = ERR_USER_BREAK Then
        result = U("\u0110\u00e3 d\u1eebng in theo y\u00eau c\u1ea7u.")
        kind = vbInformation
    Else
        result = U("Kh\u00f4ng in \u0111\u01b0\u1ee3c nh\u1eadt k\u00fd: ") & Err.Description
    End If
    Resume Cleanup
End Sub
