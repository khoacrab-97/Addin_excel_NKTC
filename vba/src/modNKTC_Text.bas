Attribute VB_Name = "modNKTC_Text"
'==============================================================================
' NHẬT KÝ TC – xử lý chữ (không phụ thuộc Excel).
' Mã nguồn viết bằng UTF-8; vba/tools/vba_ascii.py đổi chuỗi tiếng Việt sang U("\uXXXX")
' trước khi nạp vào trình soạn VBA (VBE không lưu được chữ có dấu).
'==============================================================================
Option Explicit

Private mAccents As String
Private mBases As String

' Giải mã chuỗi dạng \uXXXX (do vba_ascii.py sinh ra) thành chữ Unicode.
Public Function U(ByVal s As String) As String
    Dim i As Long, n As Long, out As String, code As Long
    n = Len(s)
    i = 1
    Do While i <= n
        If Mid$(s, i, 2) = "\u" And i + 5 <= n Then
            ' "&HXXXX" đổi sang Integer 16 bit: từ &H8000 trở lên ra số âm -> cộng 65536.
            code = CLng("&H" & UCase$(Mid$(s, i + 2, 4)))
            If code < 0 Then code = code + 65536
            out = out & ChrW(code)
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

' Hai chữ số: 5 -> "05".
Public Function Pad2(ByVal v As Long) As String
    If v < 10 Then Pad2 = "0" & CStr(v) Else Pad2 = CStr(v)
End Function

Private Sub InitAccents()
    If Len(mAccents) > 0 Then Exit Sub
    ' Mỗi nhóm chữ có dấu (chữ thường) ứng với một chữ không dấu.
    Dim groups As Variant, bases As Variant, g As Long, k As Long
    groups = Array("àáạảãâầấậẩẫăằắặẳẵ", "èéẹẻẽêềếệểễ", "ìíịỉĩ", "òóọỏõôồốộổỗơờớợởỡ", "ùúụủũưừứựửữ", "ỳýỵỷỹ", "đ")
    bases = Array("a", "e", "i", "o", "u", "y", "d")
    For g = 0 To UBound(groups)
        mAccents = mAccents & groups(g)
        For k = 1 To Len(groups(g))
            mBases = mBases & bases(g)
        Next k
    Next g
End Sub

' Chữ thường, bỏ dấu tiếng Việt.
Public Function RemoveDiacritics(ByVal s As String) As String
    Dim i As Long, ch As String, p As Long, out As String
    InitAccents
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        p = InStr(1, mAccents, ch, vbBinaryCompare)
        If p > 0 Then out = out & Mid$(mBases, p, 1) Else out = out & ch
    Next i
    RemoveDiacritics = out
End Function

' Chuẩn hóa tiêu đề để nhận diện cột: chữ thường, bỏ dấu, ký tự khác a-z/0-9 thành một khoảng trắng, bỏ đầu/cuối.
Public Function NormalizeHeader(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, gap As Boolean
    s = RemoveDiacritics(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            If gap And Len(out) > 0 Then out = out & " "
            out = out & ch
            gap = False
        Else
            gap = True
        End If
    Next i
    NormalizeHeader = out
End Function

' Chuẩn hóa tiêu đề để khớp cột theo tên (R2.17): không phân biệt hoa/thường, gộp khoảng trắng; giữ dấu.
Public Function NormalizeHeaderName(ByVal s As String) As String
    NormalizeHeaderName = LCase$(NormalizeSpaces(s))
End Function

' Giá trị ô thành chữ (giống headerTexts của bản Office.js): chuỗi giữ nguyên, số/logic thành chữ, còn lại rỗng.
Public Function CellText(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbString: CellText = v
        Case vbBoolean: If v Then CellText = "true" Else CellText = "false"
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte: CellText = CStr(v)
        Case Else: CellText = ""
    End Select
End Function

' Ô trống: Empty, chuỗi chỉ có khoảng trắng.
Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then IsBlankValue = True: Exit Function
    If VarType(v) = vbString Then IsBlankValue = (Len(Trim$(v)) = 0)
End Function

' Chữ cột Excel từ chỉ số 0-based: 0 -> "A", 27 -> "AB".
Public Function ColumnLetter(ByVal index0 As Long) As String
    Dim n As Long, out As String
    n = index0 + 1
    Do While n > 0
        out = Chr$(65 + ((n - 1) Mod 26)) & out
        n = (n - 1) \ 26
    Loop
    ColumnLetter = out
End Function

' Nhãn ngắn: gộp khoảng trắng, tối đa 60 ký tự (giống shortLabel của bản Office.js).
Public Function ShortLabel(ByVal s As String) As String
    Const MAX_LEN As Long = 60
    s = NormalizeSpaces(s)
    If Len(s) > MAX_LEN Then s = Left$(s, MAX_LEN - 1) & ChrW(8230)
    ShortLabel = s
End Function

' Gộp mọi khoảng trắng/xuống dòng thành một dấu cách, bỏ khoảng trắng hai đầu.
Public Function NormalizeSpaces(ByVal s As String) As String
    Dim parts As Variant, i As Long, out As String
    s = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
    parts = Split(Trim$(s), " ")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Len(out) > 0 Then out = out & " "
            out = out & parts(i)
        End If
    Next i
    NormalizeSpaces = out
End Function

' Dòng trong danh sách chọn cột: "Cột C – Nội dung công việc". sheetColumn0: chỉ số cột 0-based trên sheet.
Public Function ColumnOptionLabel(ByVal header As Variant, ByVal sheetColumn0 As Long) As String
    Dim label As String
    label = ShortLabel(CellText(header))
    If Len(label) = 0 Then label = "(không có tiêu đề)"
    ColumnOptionLabel = "Cột " & ColumnLetter(sheetColumn0) & " – " & label
End Function
