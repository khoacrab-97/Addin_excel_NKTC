Attribute VB_Name = "modNKTC_Text"
'==============================================================================
' NHAT KY TC ? xu ly chu (khong phu thuoc Excel).
' Ma nguon viet bang UTF-8; vba/tools/vba_ascii.py doi chuoi tieng Viet sang U("\uXXXX")
' truoc khi nap vao trinh soan VBA (VBE khong luu duoc chu co dau).
'==============================================================================
Option Explicit

Private mAccents As String
Private mBases As String

' Giai ma chuoi dang \uXXXX (do vba_ascii.py sinh ra) thanh chu Unicode.
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

' Hai chu so: 5 -> "05".
Public Function Pad2(ByVal v As Long) As String
    If v < 10 Then Pad2 = "0" & CStr(v) Else Pad2 = CStr(v)
End Function

Private Sub InitAccents()
    If Len(mAccents) > 0 Then Exit Sub
    ' Moi nhom chu co dau (chu thuong) ung voi mot chu khong dau.
    Dim groups As Variant, bases As Variant, g As Long, k As Long
    groups = Array(U("\u00e0\u00e1\u1ea1\u1ea3\u00e3\u00e2\u1ea7\u1ea5\u1ead\u1ea9\u1eab\u0103\u1eb1\u1eaf\u1eb7\u1eb3\u1eb5"), U("\u00e8\u00e9\u1eb9\u1ebb\u1ebd\u00ea\u1ec1\u1ebf\u1ec7\u1ec3\u1ec5"), U("\u00ec\u00ed\u1ecb\u1ec9\u0129"), U("\u00f2\u00f3\u1ecd\u1ecf\u00f5\u00f4\u1ed3\u1ed1\u1ed9\u1ed5\u1ed7\u01a1\u1edd\u1edb\u1ee3\u1edf\u1ee1"), U("\u00f9\u00fa\u1ee5\u1ee7\u0169\u01b0\u1eeb\u1ee9\u1ef1\u1eed\u1eef"), U("\u1ef3\u00fd\u1ef5\u1ef7\u1ef9"), U("\u0111"))
    bases = Array("a", "e", "i", "o", "u", "y", "d")
    For g = 0 To UBound(groups)
        mAccents = mAccents & groups(g)
        For k = 1 To Len(groups(g))
            mBases = mBases & bases(g)
        Next k
    Next g
End Sub

' Chu thuong, bo dau tieng Viet.
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

' Chuan hoa tieu de de nhan dien cot: chu thuong, bo dau, ky tu khac a-z/0-9 thanh mot khoang trang, bo dau/cuoi.
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

' Chuan hoa tieu de de khop cot theo ten (R2.17): khong phan biet hoa/thuong, gop khoang trang; giu dau.
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

' Gia tri o thanh chu (giong headerTexts cua ban Office.js): chuoi giu nguyen, so/logic thanh chu, con lai rong.
Public Function CellText(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbString: CellText = v
        Case vbBoolean: If v Then CellText = "true" Else CellText = "false"
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte: CellText = CStr(v)
        Case Else: CellText = ""
    End Select
End Function

' O trong: Empty, chuoi chi co khoang trang.
Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then IsBlankValue = True: Exit Function
    If VarType(v) = vbString Then IsBlankValue = (Len(Trim$(v)) = 0)
End Function
