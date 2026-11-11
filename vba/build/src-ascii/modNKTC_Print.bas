Attribute VB_Name = "modNKTC_Print"
'==============================================================================
' NHAT KY TC ? In nhat ky: chon cac ngay can in, ten file PDF, ten sheet tam (khong phu thuoc Excel).
'==============================================================================
Option Explicit

Public Const PRINT_CURRENT As Long = 1   ' ngay dang hien thi
Public Const PRINT_ONE As Long = 2       ' mot ngay
Public Const PRINT_RANGE As Long = 3     ' khoang ngay
Public Const PRINT_ALL As Long = 4       ' tat ca
Public Const PRINT_CONFIRM_DAYS As Long = 50   ' in nhieu hon so ngay nay -> hoi xac nhan

' Chon cac ngay can in theo che do. days(): danh sach ngay tang dan (BuildDayIndex).
' fromDay/toDay: so ngay (DayNum) - PRINT_ONE dung fromDay; PRINT_RANGE dung ca hai (doi cho neu from > to);
' PRINT_CURRENT dung currentDay (0 neu chua hien thi ngay nao).
' Ket qua: picked() (0-based) la CHI SO trong days(); pickedCount. Tra ve chuoi loi (rong neu duoc).
Public Function SelectPrintDays(ByRef days() As NKTC_SourceDay, ByVal dayCount As Long, ByVal mode As Long, _
    ByVal fromDay As Long, ByVal toDay As Long, ByVal currentDay As Long, _
    ByRef picked() As Long, ByRef pickedCount As Long) As String
    Dim idx As Long, i As Long, tmp As Long
    pickedCount = 0
    ReDim picked(0 To 0)
    If dayCount <= 0 Then
        SelectPrintDays = U("V\u00f9ng ngu\u1ed3n kh\u00f4ng c\u00f3 ng\u00e0y h\u1ee3p l\u1ec7 \u0111\u1ec3 in.")
        Exit Function
    End If
    ReDim picked(0 To dayCount - 1)
    Select Case mode
        Case PRINT_CURRENT
            idx = -1
            If currentDay > 0 Then idx = IndexOfDay(days, dayCount, currentDay)
            If idx = -1 Then
                If currentDay > 0 Then
                    SelectPrintDays = U("Ng\u00e0y \u0111ang hi\u1ec3n th\u1ecb ") & FormatDay(currentDay) & U(" kh\u00f4ng c\u00f2n trong v\u00f9ng ngu\u1ed3n.")
                Else
                    SelectPrintDays = U("Bi\u1ec3u m\u1eabu ch\u01b0a hi\u1ec3n th\u1ecb ng\u00e0y n\u00e0o. B\u1ea5m m\u1ed9t n\u00fat Chuy\u1ec3n ng\u00e0y tr\u01b0\u1edbc, ho\u1eb7c ch\u1ecdn ng\u00e0y c\u1ea7n in.")
                End If
                Exit Function
            End If
            picked(0) = idx: pickedCount = 1
        Case PRINT_ONE
            If fromDay = 0 Then
                SelectPrintDays = U("Nh\u1eadp ng\u00e0y c\u1ea7n in (dd/mm/yyyy).")
                Exit Function
            End If
            idx = IndexOfDay(days, dayCount, fromDay)
            If idx = -1 Then
                SelectPrintDays = U("Kh\u00f4ng c\u00f3 ng\u00e0y ") & FormatDay(fromDay) & U(" trong v\u00f9ng ngu\u1ed3n.")
                Exit Function
            End If
            picked(0) = idx: pickedCount = 1
        Case PRINT_RANGE
            If fromDay = 0 Or toDay = 0 Then
                SelectPrintDays = U("Nh\u1eadp \u0111\u1ee7 T\u1eeb ng\u00e0y v\u00e0 \u0110\u1ebfn ng\u00e0y (dd/mm/yyyy).")
                Exit Function
            End If
            If fromDay > toDay Then tmp = fromDay: fromDay = toDay: toDay = tmp
            For i = 0 To dayCount - 1
                If days(i).DayNum >= fromDay And days(i).DayNum <= toDay Then
                    picked(pickedCount) = i
                    pickedCount = pickedCount + 1
                End If
            Next i
            If pickedCount = 0 Then
                SelectPrintDays = U("Kh\u00f4ng c\u00f3 ng\u00e0y n\u00e0o t\u1eeb ") & FormatDay(fromDay) & U(" \u0111\u1ebfn ") & FormatDay(toDay) & U(" trong v\u00f9ng ngu\u1ed3n.")
                Exit Function
            End If
        Case PRINT_ALL
            For i = 0 To dayCount - 1
                picked(i) = i
            Next i
            pickedCount = dayCount
        Case Else
            SelectPrintDays = U("Ch\u1ebf \u0111\u1ed9 in kh\u00f4ng h\u1ee3p l\u1ec7.")
    End Select
End Function

' Ngay dang dd-mm-yyyy (dung trong ten file / ten sheet).
Private Function DashDay(ByVal dayNum As Long) As String
    DashDay = Replace(FormatDay(dayNum), "/", "-")
End Function

' Ten file PDF goi y: mot ngay -> "NKTC_15-06-2026.pdf"; nhieu ngay -> "NKTC_15-06-2026_den_31-12-2026.pdf".
Public Function PrintPdfFileName(ByVal firstDay As Long, ByVal lastDay As Long) As String
    If firstDay = lastDay Then
        PrintPdfFileName = "NKTC_" & DashDay(firstDay) & ".pdf"
    Else
        PrintPdfFileName = "NKTC_" & DashDay(firstDay) & "_den_" & DashDay(lastDay) & ".pdf"
    End If
End Function

' Ten sheet tam cho mot ngay: "Ngay_15-06-2026" (<= 31 ky tu, khong ky tu cam).
Public Function PrintSheetName(ByVal dayNum As Long) As String
    PrintSheetName = "Ngay_" & DashDay(dayNum)
End Function
