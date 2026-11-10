Attribute VB_Name = "modNKTC_Update"
'==============================================================================
' NHẬT KÝ TC – tự cập nhật (theo mô hình TienBangChu).
' 1. Đọc version.txt ở kênh cập nhật (version=, file=, sha256=, size=, notes=).
' 2. Có bản mới -> tải NhatKyThiCong_Core.xlam, kiểm tra kích thước, chữ ký ZIP ("PK"), SHA-256.
' 3. Đạt -> lưu NhatKyThiCong_Core.pending + .pending.ok (dòng 1: phiên bản, dòng 2: SHA-256) cạnh add-in.
' 4. Lần mở Excel sau, bộ nạp NhatKyThiCong.xlam kiểm lại SHA-256 rồi thay tệp lõi.
' Tự kiểm tra mỗi ngày một lần (30 giây sau khi mở Excel, im lặng nếu không có gì mới); nút "Kiểm tra cập nhật".
'==============================================================================
Option Explicit

' Kênh cập nhật (GitHub Pages của repo), có "/" cuối. settings.txt có thể đổi (UpdateUrl=, chỉ https hoặc loopback).
Private Const UPDATE_BASE_URL As String = "https://khoacrab-97.github.io/Addin_excel_NKTC/vba/"
Public Const CORE_FILE As String = "NhatKyThiCong_Core.xlam"
Public Const PENDING_FILE As String = "NhatKyThiCong_Core.pending"
Public Const READY_FILE As String = "NhatKyThiCong_Core.pending.ok"
Private Const SETTINGS_FILE As String = "settings.txt"
Private Const MAX_CORE_SIZE As Long = 20000000

Public Const UPD_ERROR As Long = -1
Public Const UPD_LATEST As Long = 0
Public Const UPD_DOWNLOADED As Long = 1
Public Const UPD_PENDING As Long = 2
Public Const UPD_NOT_CONFIGURED As Long = 3

#If VBA7 Then
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

Private mLoaded As Boolean
Private mLastCheck As String
Private mUpdateUrl As String

'------------------------------------------------------------------------------
' Phần thuần (có kiểm thử): đọc version.txt, so sánh phiên bản, kiểm tra địa chỉ.
'------------------------------------------------------------------------------
' version.txt: mỗi dòng "khóa=giá trị". Tên tệp không được chứa đường dẫn (/, \, :, ..).
Public Sub ParseManifest(ByVal s As String, ByRef ver As String, ByRef fileName As String, ByRef sha As String, _
                         ByRef size As Long, ByRef notes As String)
    Dim ln As Variant, p As Long, k As String, v As String
    ver = "": fileName = "": sha = "": size = 0: notes = ""
    s = Replace(s, vbCr, "")
    For Each ln In Split(s, vbLf)
        p = InStr(ln, "=")
        If p > 1 Then
            k = LCase$(Trim$(Left$(ln, p - 1)))
            v = Trim$(Mid$(ln, p + 1))
            Select Case k
                Case "version": If IsVersion(v) Then ver = v
                Case "file": If v Like "*[\/:]*" Or InStr(v, "..") > 0 Then fileName = "" Else fileName = v
                Case "sha256": If IsSha256(v) Then sha = LCase$(v)
                Case "size": If IsNumeric(v) And Len(v) <= 9 Then size = CLng(v)
                Case "notes": notes = Left$(v, 500)
            End Select
        End If
    Next ln
End Sub

' Phiên bản dạng số chấm số: 3.0.0, 3.1.12 (tối đa 4 phần).
Public Function IsVersion(ByVal v As String) As Boolean
    Dim parts As Variant, i As Long, j As Long
    If Len(v) = 0 Or Len(v) > 23 Then Exit Function
    parts = Split(v, ".")
    If UBound(parts) < 1 Or UBound(parts) > 3 Then Exit Function
    For i = 0 To UBound(parts)
        If Len(parts(i)) = 0 Or Len(parts(i)) > 5 Then Exit Function
        For j = 1 To Len(parts(i))
            If Mid$(parts(i), j, 1) < "0" Or Mid$(parts(i), j, 1) > "9" Then Exit Function
        Next j
    Next i
    IsVersion = True
End Function

Public Function IsSha256(ByVal v As String) As Boolean
    Dim i As Long, ch As String
    If Len(v) <> 64 Then Exit Function
    v = LCase$(v)
    For i = 1 To 64
        ch = Mid$(v, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "f")) Then Exit Function
    Next i
    IsSha256 = True
End Function

' 1 nếu a > b, 0 nếu bằng, -1 nếu a < b (so từng phần số: 3.10.0 > 3.9.0; phần không phải số = 0).
Public Function CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim pa As Variant, pb As Variant, i As Long, x As Long, y As Long
    pa = Split(a, ".")
    pb = Split(b, ".")
    For i = 0 To 3
        x = 0: y = 0
        If i <= UBound(pa) Then x = VersionPart(pa(i))
        If i <= UBound(pb) Then y = VersionPart(pb(i))
        If x > y Then CompareVersions = 1: Exit Function
        If x < y Then CompareVersions = -1: Exit Function
    Next i
End Function

Private Function VersionPart(ByVal s As String) As Long
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 5 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    VersionPart = CLng(s)
End Function

' Địa chỉ cập nhật thay thế (settings.txt): chỉ https, hoặc http tới máy này (kiểm thử); không có thông tin đăng nhập.
Public Function IsAllowedUpdateUrl(ByVal url As String) As Boolean
    Dim l As String
    l = LCase$(Trim$(url))
    If InStr(l, "@") > 0 Or InStr(l, " ") > 0 Then Exit Function
    If Left$(l, 8) = "https://" Then IsAllowedUpdateUrl = (Len(l) > 8): Exit Function
    IsAllowedUpdateUrl = (l Like "http://127.0.0.1[:/]*" Or l Like "http://localhost[:/]*")
End Function

'------------------------------------------------------------------------------
' Thiết lập (settings.txt cạnh add-in): LastCheck=, UpdateUrl=
'------------------------------------------------------------------------------
Private Function FileInAddinFolder(ByVal name As String) As String
    FileInAddinFolder = ThisWorkbook.Path & Application.PathSeparator & name
End Function

Private Sub LoadSettings()
    Dim text As String, ln As Variant, p As Long
    If mLoaded Then Exit Sub
    mLoaded = True
    text = ReadTextFile(FileInAddinFolder(SETTINGS_FILE))
    For Each ln In Split(Replace(text, vbCr, ""), vbLf)
        p = InStr(ln, "=")
        If p > 1 Then
            Select Case LCase$(Trim$(Left$(ln, p - 1)))
                Case "lastcheck": mLastCheck = Trim$(Mid$(ln, p + 1))
                Case "updateurl": mUpdateUrl = Trim$(Mid$(ln, p + 1))
            End Select
        End If
    Next ln
End Sub

Private Sub SaveSettings()
    Dim text As String
    text = "LastCheck=" & mLastCheck & vbCrLf
    If Len(mUpdateUrl) > 0 Then text = text & "UpdateUrl=" & mUpdateUrl & vbCrLf
    On Error Resume Next
    WriteTextFile FileInAddinFolder(SETTINGS_FILE), text
End Sub

Private Function UpdateBase() As String
    Dim s As String
    LoadSettings
    s = UPDATE_BASE_URL
    If Len(mUpdateUrl) > 0 Then
        If IsAllowedUpdateUrl(mUpdateUrl) Then s = mUpdateUrl
    End If
    If Right$(s, 1) <> "/" Then s = s & "/"
    UpdateBase = s
End Function

'------------------------------------------------------------------------------
' Lịch kiểm tra và lệnh
'------------------------------------------------------------------------------
' Gọi khi lõi được mở: nếu đã quá 1 ngày từ lần kiểm tra trước thì kiểm tra sau 30 giây (im lặng).
Public Sub ScheduleUpdateCheck()
    Dim last As Date
    On Error Resume Next
    ' Excel do chương trình khác điều khiển (build, kiểm thử) -> không hẹn giờ.
    If Not Application.UserControl Then Exit Sub
    LoadSettings
    If Len(mLastCheck) > 0 Then last = CDate(Replace(mLastCheck, "T", " "))
    If last = 0 Or Now - last >= 1 Then
        Application.OnTime Now + TimeSerial(0, 0, 30), "'" & ThisWorkbook.Name & "'!modNKTC_Update.AutoUpdate"
    End If
End Sub

Public Sub AutoUpdate()
    On Error Resume Next
    CheckForUpdate False
End Sub

' Nút "Kiểm tra cập nhật" (showResult = True) hoặc tự động (False: chỉ báo khi vừa tải bản mới).
Public Sub CheckForUpdate(Optional ByVal showResult As Boolean = True)
    Dim result As Long, message As String
    result = RunUpdate(message)
    Select Case result
        Case UPD_DOWNLOADED: ShowInfo message
        Case UPD_ERROR: If showResult Then ShowInfo message, vbExclamation
        Case Else: If showResult Then ShowInfo message
    End Select
End Sub

' Dùng cho kiểm thử tự động (Application.Run): "mã|thông điệp", không hiện hộp thoại.
Public Function NKTC_UpdateSilent() As String
    Dim message As String, result As Long
    result = RunUpdate(message)
    NKTC_UpdateSilent = result & "|" & message
End Function

Private Function RunUpdate(ByRef message As String) As Long
    Dim base As String, body() As Byte, status As Long, failure As String
    Dim ver As String, fileName As String, sha As String, size As Long, notes As String
    On Error GoTo Loi
    RunUpdate = UPD_ERROR
    base = UpdateBase()

    If Not Download(base & "version.txt", body, status, failure) Then GoTo KhongDuoc
    ParseManifest BytesToText(body), ver, fileName, sha, size, notes
    mLastCheck = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SaveSettings
    If Len(ver) = 0 Or Len(fileName) = 0 Or Len(sha) = 0 Then failure = "Tệp version.txt không hợp lệ.": GoTo KhongDuoc

    If CompareVersions(ver, APP_VERSION) <= 0 Then
        message = "Bạn đang dùng phiên bản mới nhất (" & APP_VERSION & ")."
        RunUpdate = UPD_LATEST
        Exit Function
    End If
    If ReadyVersion() = ver Then
        message = "Bản cập nhật " & ver & " đã được tải về. Hãy đóng và mở lại Excel để dùng bản mới."
        RunUpdate = UPD_PENDING
        Exit Function
    End If

    If Not Download(base & fileName, body, status, failure) Then GoTo KhongDuoc
    If UBound(body) < 4 Then failure = "Tệp tải về rỗng.": GoTo KhongDuoc
    If UBound(body) + 1 > MAX_CORE_SIZE Then failure = "Tệp tải về quá lớn.": GoTo KhongDuoc
    If size > 0 And UBound(body) + 1 <> size Then failure = "Kích thước tệp tải về không khớp.": GoTo KhongDuoc
    If body(0) <> 80 Or body(1) <> 75 Then failure = "Tệp tải về không phải add-in Excel.": GoTo KhongDuoc
    If Sha256Hex(body) <> sha Then failure = "Mã kiểm tra SHA-256 không khớp, đã hủy bản tải về.": GoTo KhongDuoc

    DeleteIfExists FileInAddinFolder(READY_FILE)
    WriteBytes FileInAddinFolder(PENDING_FILE), body
    WriteTextFile FileInAddinFolder(READY_FILE), ver & vbCrLf & sha & vbCrLf

    message = "Đã tải bản cập nhật Nhật ký TC " & ver & " (đang dùng " & APP_VERSION & ")." & vbLf & _
        "Hãy đóng và mở lại Excel để dùng bản mới." & IIf(Len(notes) > 0, vbLf & vbLf & notes, "")
    RunUpdate = UPD_DOWNLOADED
    Exit Function
KhongDuoc:
    message = "Không kiểm tra được bản cập nhật." & vbLf & failure
    Exit Function
Loi:
    message = "Không kiểm tra được bản cập nhật: " & Err.Description
    RunUpdate = UPD_ERROR
End Function

' Phiên bản trong .pending.ok nếu bản tải về còn đó (chờ mở lại Excel); rỗng nếu không có.
Private Function ReadyVersion() As String
    Dim text As String
    If Len(Dir$(FileInAddinFolder(PENDING_FILE))) = 0 Then Exit Function
    text = ReadTextFile(FileInAddinFolder(READY_FILE))
    ReadyVersion = Trim$(Split(Replace(text, vbCr, "") & vbLf, vbLf)(0))
End Function

'------------------------------------------------------------------------------
' Tải về, tệp, SHA-256
'------------------------------------------------------------------------------
Private Function Download(ByVal url As String, ByRef body() As Byte, ByRef status As Long, ByRef failure As String) As Boolean
    Dim h As Object
    On Error GoTo Loi
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 5000, 5000, 15000, 60000
    On Error Resume Next
    h.Option(9) = &H2800                                    ' TLS 1.2 / 1.3
    If Err.Number <> 0 Then Err.Clear: h.Option(9) = &H800  ' WinHttp chưa có TLS 1.3: chỉ TLS 1.2
    On Error GoTo Loi
    h.Open "GET", url, False
    h.SetRequestHeader "User-Agent", "NhatKyTC/" & APP_VERSION
    h.SetRequestHeader "Cache-Control", "no-cache"
    h.Send
    status = h.Status
    If status <> 200 Then failure = "HTTP " & status & " - " & url: Exit Function
    body = h.ResponseBody
    Download = True
    Exit Function
Loi:
    failure = Err.Description
End Function

Private Function BytesToText(ByRef b() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    BytesToText = st.ReadText
    st.Close
End Function

Private Sub WriteBytes(ByVal path As String, ByRef b() As Byte)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write b
    st.SaveToFile path, 2
    st.Close
End Sub

Private Sub WriteTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer
    On Error GoTo Het
    If Len(Dir$(path)) = 0 Then Exit Function
    f = FreeFile
    Open path For Input As #f
    ReadTextFile = Input$(LOF(f), f)
    Close #f
    Exit Function
Het:
    On Error Resume Next
    Close #f
End Function

Private Sub DeleteIfExists(ByVal path As String)
    If Len(Dir$(path)) > 0 Then Kill path
End Sub

Public Function Sha256Hex(ByRef b() As Byte) As String
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
