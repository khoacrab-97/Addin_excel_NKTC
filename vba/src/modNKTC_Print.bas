Attribute VB_Name = "modNKTC_Print"
'==============================================================================
' NHẬT KÝ TC – In nhật ký: chọn các ngày cần in, tên file PDF, tên sheet tạm (không phụ thuộc Excel).
'==============================================================================
Option Explicit

Public Const PRINT_CURRENT As Long = 1   ' ngày đang hiển thị
Public Const PRINT_ONE As Long = 2       ' một ngày
Public Const PRINT_RANGE As Long = 3     ' khoảng ngày
Public Const PRINT_ALL As Long = 4       ' tất cả
Public Const PRINT_CONFIRM_DAYS As Long = 50   ' in nhiều hơn số ngày này -> hỏi xác nhận

' Chọn các ngày cần in theo chế độ. days(): danh sách ngày tăng dần (BuildDayIndex).
' fromDay/toDay: số ngày (DayNum) - PRINT_ONE dùng fromDay; PRINT_RANGE dùng cả hai (đổi chỗ nếu from > to);
' PRINT_CURRENT dùng currentDay (0 nếu chưa hiển thị ngày nào).
' Kết quả: picked() (0-based) là CHỈ SỐ trong days(); pickedCount. Trả về chuỗi lỗi (rỗng nếu được).
Public Function SelectPrintDays(ByRef days() As NKTC_SourceDay, ByVal dayCount As Long, ByVal mode As Long, _
    ByVal fromDay As Long, ByVal toDay As Long, ByVal currentDay As Long, _
    ByRef picked() As Long, ByRef pickedCount As Long) As String
    Dim idx As Long, i As Long, tmp As Long
    pickedCount = 0
    ReDim picked(0 To 0)
    If dayCount <= 0 Then
        SelectPrintDays = "Vùng nguồn không có ngày hợp lệ để in."
        Exit Function
    End If
    ReDim picked(0 To dayCount - 1)
    Select Case mode
        Case PRINT_CURRENT
            idx = -1
            If currentDay > 0 Then idx = IndexOfDay(days, dayCount, currentDay)
            If idx = -1 Then
                If currentDay > 0 Then
                    SelectPrintDays = "Ngày đang hiển thị " & FormatDay(currentDay) & " không còn trong vùng nguồn."
                Else
                    SelectPrintDays = "Biểu mẫu chưa hiển thị ngày nào. Bấm một nút Chuyển ngày trước, hoặc chọn ngày cần in."
                End If
                Exit Function
            End If
            picked(0) = idx: pickedCount = 1
        Case PRINT_ONE
            If fromDay = 0 Then
                SelectPrintDays = "Nhập ngày cần in (dd/mm/yyyy)."
                Exit Function
            End If
            idx = IndexOfDay(days, dayCount, fromDay)
            If idx = -1 Then
                SelectPrintDays = "Không có ngày " & FormatDay(fromDay) & " trong vùng nguồn."
                Exit Function
            End If
            picked(0) = idx: pickedCount = 1
        Case PRINT_RANGE
            If fromDay = 0 Or toDay = 0 Then
                SelectPrintDays = "Nhập đủ Từ ngày và Đến ngày (dd/mm/yyyy)."
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
                SelectPrintDays = "Không có ngày nào từ " & FormatDay(fromDay) & " đến " & FormatDay(toDay) & " trong vùng nguồn."
                Exit Function
            End If
        Case PRINT_ALL
            For i = 0 To dayCount - 1
                picked(i) = i
            Next i
            pickedCount = dayCount
        Case Else
            SelectPrintDays = "Chế độ in không hợp lệ."
    End Select
End Function

' Ngày dạng dd-mm-yyyy (dùng trong tên file / tên sheet).
Private Function DashDay(ByVal dayNum As Long) As String
    DashDay = Replace(FormatDay(dayNum), "/", "-")
End Function

' Tên file PDF gợi ý: một ngày -> "NKTC_15-06-2026.pdf"; nhiều ngày -> "NKTC_15-06-2026_den_31-12-2026.pdf".
Public Function PrintPdfFileName(ByVal firstDay As Long, ByVal lastDay As Long) As String
    If firstDay = lastDay Then
        PrintPdfFileName = "NKTC_" & DashDay(firstDay) & ".pdf"
    Else
        PrintPdfFileName = "NKTC_" & DashDay(firstDay) & "_den_" & DashDay(lastDay) & ".pdf"
    End If
End Function

' Tên sheet tạm cho một ngày: "Ngay_15-06-2026" (<= 31 ký tự, không ký tự cấm).
Public Function PrintSheetName(ByVal dayNum As Long) As String
    PrintSheetName = "Ngay_" & DashDay(dayNum)
End Function
