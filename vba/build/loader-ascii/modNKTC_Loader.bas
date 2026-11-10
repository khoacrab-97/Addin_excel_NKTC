Attribute VB_Name = "modNKTC_Loader"
'==============================================================================
' NHAT KY TC ? bo nap (NhatKyThiCong.xlam). Excel mo bo nap khi khoi dong (bo cai dang ky khoa OPEN).
' 1. Co ban cap nhat da tai (NhatKyThiCong_Core.pending + .pending.ok) -> kiem lai SHA-256 roi thay tep loi
'    (giu ban cu o .bak; loi thi khoi phuc).
' 2. Mo NhatKyThiCong_Core.xlam (tab Ribbon "Nhat ky TC", chuc nang, tu cap nhat).
' Bo nap khong tu cap nhat: loi dang mo khong ghi de duoc chinh no nen viec thay tep lam o day.
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

' Mo loi; tra ve ly do neu khong mo duoc (rong = da mo). Build goi truc tiep qua Application.Run de kiem thu.
Public Function NKTC_TryLoadCore() As String
    Dim folder As String, wb As Workbook
    On Error GoTo Loi
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(CORE_FILE) Then Exit Function
    Next wb
    folder = ThisWorkbook.Path & Application.PathSeparator
    ApplyPendingUpdate folder
    If Len(Dir$(folder & CORE_FILE)) = 0 Then
        NKTC_TryLoadCore = U("Kh\u00f4ng t\u00ecm th\u1ea5y t\u1ec7p ") & CORE_FILE & " trong " & ThisWorkbook.Path & "." & vbLf & _
            U("H\u00e3y ch\u1ea1y l\u1ea1i b\u1ed9 c\u00e0i Nh\u1eadt k\u00fd TC.")
        Exit Function
    End If
    Application.Workbooks.Open Filename:=folder & CORE_FILE, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False
    Exit Function
Loi:
    NKTC_TryLoadCore = U("Kh\u00f4ng m\u1edf \u0111\u01b0\u1ee3c Nh\u1eadt k\u00fd TC: ") & Err.Description
End Function

' Thay tep loi bang ban cap nhat da tai, neu ban do con nguyen ven (SHA-256 khop .pending.ok).
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
    ' Khong thay duoc (tep dang bi khoa...) -> tra lai ban cu, lan sau thu lai.
    On Error Resume Next
    If movedOld And Len(Dir$(core)) = 0 Then Name backup As core
    Exit Sub
Hong:
    ' Ban tai ve hong hoac khong khop -> bo; loi se tai lai khi kiem tra cap nhat.
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
    title = U("Nh\u1eadt k\u00fd TC")
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(title), vbExclamation Or &H10000
End Sub

' Giai ma chuoi \uXXXX (do vba/tools/vba_ascii.py sinh ra) ? bo nap co ban rieng vi chay khi chua co loi.
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
