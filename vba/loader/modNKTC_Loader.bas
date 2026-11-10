Attribute VB_Name = "modNKTC_Loader"
'==============================================================================
' NHẬT KÝ TC – bộ nạp (NhatKyThiCong.xlam). Excel mở bộ nạp khi khởi động (bộ cài đăng ký khóa OPEN).
' 1. Có bản cập nhật đã tải (NhatKyThiCong_Core.pending + .pending.ok) -> kiểm lại SHA-256 rồi thay tệp lõi
'    (giữ bản cũ ở .bak; lỗi thì khôi phục).
' 2. Mở NhatKyThiCong_Core.xlam (tab Ribbon "Nhật ký TC", chức năng, tự cập nhật).
' Bộ nạp không tự cập nhật: lõi đang mở không ghi đè được chính nó nên việc thay tệp làm ở đây.
'==============================================================================
Option Explicit

Public Const LOADER_VERSION As String = "1.0.0"
Private Const CORE_FILE As String = "NhatKyThiCong_Core.xlam"
Private Const PENDING_FILE As String = "NhatKyThiCong_Core.pending"
Private Const READY_FILE As String = "NhatKyThiCong_Core.pending.ok"
Private Const BACKUP_FILE As String = "NhatKyThiCong_Core.bak"

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, _
    ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal hProv As LongPtr, ByVal Algid As Long, _
    ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hHash As LongPtr, ByRef pbData As Byte, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hHash As LongPtr, ByVal dwParam As Long, _
    ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function CryptAcquireContextW Lib "advapi32" (ByRef phProv As Long, ByVal pszContainer As Long, _
    ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal Algid As Long, _
    ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByRef pbData As Byte, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, _
    ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

Public Sub NKTC_LoadCore()
    Dim msg As String
    msg = NKTC_TryLoadCore()
    If Len(msg) > 0 Then Notify msg
End Sub

' Mở lõi; trả về lý do nếu không mở được (rỗng = đã mở). Build gọi trực tiếp qua Application.Run để kiểm thử.
Public Function NKTC_TryLoadCore() As String
    Dim folder As String
    If IsCoreOpen() Then Exit Function
    On Error GoTo Loi
    folder = ThisWorkbook.Path & Application.PathSeparator
    ApplyPendingUpdate folder
    If Len(Dir$(folder & CORE_FILE)) = 0 Then
        NKTC_TryLoadCore = "Không tìm thấy tệp " & CORE_FILE & " trong " & ThisWorkbook.Path & "." & vbLf & _
            "Hãy chạy lại bộ cài Nhật ký TC."
        Exit Function
    End If
    Application.Workbooks.Open Filename:=folder & CORE_FILE, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False
    Exit Function
Loi:
    NKTC_TryLoadCore = "Không mở được Nhật ký TC: " & Err.Description
End Function

' Add-in (.xlam) đang mở không có trong danh sách khi duyệt Application.Workbooks -> tìm theo tên.
Private Function IsCoreOpen() As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks(CORE_FILE)
    IsCoreOpen = Not wb Is Nothing
End Function

' Thay tệp lõi bằng bản cập nhật đã tải, nếu bản đó còn nguyên vẹn (SHA-256 khớp .pending.ok).
Private Sub ApplyPendingUpdate(ByVal folder As String)
    Dim pending As String, ready As String, core As String, backup As String, info As String
    Dim lines As Variant, expected As String, body() As Byte, movedOld As Boolean
    pending = folder & PENDING_FILE
    ready = folder & READY_FILE
    core = folder & CORE_FILE
    backup = folder & BACKUP_FILE
    If Len(Dir$(pending)) = 0 Then
        If Len(Dir$(ready)) > 0 Then SafeKill ready
        Exit Sub
    End If
    If Len(Dir$(ready)) = 0 Then SafeKill pending: Exit Sub

    info = ReadTextFile(ready)
    lines = Split(Replace(info, vbCr, "") & vbLf & vbLf, vbLf)
    expected = LCase$(Trim$(lines(1)))
    body = ReadBytes(pending)
    If Len(expected) <> 64 Or UBound(body) < 4 Then GoTo Hong
    If body(0) <> 80 Or body(1) <> 75 Then GoTo Hong
    If Sha256Hex(body) <> expected Then GoTo Hong

    On Error GoTo KhoiPhuc
    If Len(Dir$(backup)) > 0 Then Kill backup
    If Len(Dir$(core)) > 0 Then Name core As backup: movedOld = True
    Name pending As core
    SafeKill ready
    Exit Sub
KhoiPhuc:
    ' Không thay được (tệp đang bị khóa...) -> trả lại bản cũ, lần sau thử lại.
    On Error Resume Next
    If movedOld And Len(Dir$(core)) = 0 Then Name backup As core
    Exit Sub
Hong:
    ' Bản tải về hỏng hoặc không khớp -> bỏ; lõi sẽ tải lại khi kiểm tra cập nhật.
    SafeKill pending
    SafeKill ready
End Sub

Private Sub SafeKill(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer
    On Error GoTo Het
    f = FreeFile
    Open path For Input As #f
    ReadTextFile = Input$(LOF(f), f)
    Close #f
    Exit Function
Het:
    On Error Resume Next
    Close #f
End Function

Private Function ReadBytes(ByVal path As String) As Byte()
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    Else
        ReDim b(0 To 0)
    End If
    Close #f
    ReadBytes = b
End Function

Private Function Sha256Hex(ByRef b() As Byte) As String
#If VBA7 Then
    Dim hProv As LongPtr, hHash As LongPtr
#Else
    Dim hProv As Long, hHash As Long
#End If
    Dim hash(0 To 31) As Byte, n As Long, i As Long, s As String
    If CryptAcquireContextW(hProv, 0, 0, 24, &HF0000000) = 0 Then Exit Function   ' PROV_RSA_AES, CRYPT_VERIFYCONTEXT
    If CryptCreateHash(hProv, &H800C&, 0, 0, hHash) <> 0 Then                      ' CALG_SHA_256
        If CryptHashData(hHash, b(LBound(b)), UBound(b) - LBound(b) + 1, 0) <> 0 Then
            n = 32
            If CryptGetHashParam(hHash, 2, hash(0), n, 0) <> 0 Then                ' HP_HASHVAL
                For i = 0 To 31
                    s = s & Right$("0" & LCase$(Hex$(hash(i))), 2)
                Next i
            End If
        End If
        CryptDestroyHash hHash
    End If
    CryptReleaseContext hProv, 0
    Sha256Hex = s
End Function

Private Sub Notify(ByVal text As String)
    Dim title As String
    title = "Nhật ký TC"
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(title), vbExclamation Or &H10000
End Sub

' Giải mã chuỗi \uXXXX (do vba/tools/vba_ascii.py sinh ra) — bộ nạp có bản riêng vì chạy khi chưa có lõi.
Private Function U(ByVal s As String) As String
    Dim i As Long, n As Long, out As String, code As Long
    n = Len(s)
    i = 1
    Do While i <= n
        If Mid$(s, i, 2) = "\u" And i + 5 <= n Then
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
