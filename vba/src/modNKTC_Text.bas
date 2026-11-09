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
    Dim i As Long, n As Long, out As String, hex4 As String
    n = Len(s)
    i = 1
    Do While i <= n
        If Mid$(s, i, 2) = "\u" And i + 5 <= n Then
            hex4 = Mid$(s, i + 2, 4)
            out = out & ChrW(CLng("&H" & hex4 & "&"))
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
    Dim parts As Variant, i As Long, out As String
    s = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
    parts = Split(Trim$(s), " ")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Len(out) > 0 Then out = out & " "
            out = out & parts(i)
        End If
    Next i
    NormalizeHeaderName = LCase$(out)
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
