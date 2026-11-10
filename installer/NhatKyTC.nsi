; Bộ cài add-in VBA "Nhật ký TC" cho Excel 2016 / 2019 / 2021 / Microsoft 365 trên Windows.
; - Không cần quyền quản trị: chỉ ghi vào %LOCALAPPDATA%\NhatKyTC và registry HKCU.
; - Chép bộ nạp NhatKyThiCong.xlam + lõi NhatKyThiCong_Core.xlam; đăng ký bộ nạp với Excel bằng giá trị OPEN, OPEN1, …
;   trong HKCU\Software\Microsoft\Office\16.0\Excel\Options (cơ chế Excel dùng cho add-in tự mở khi khởi động),
;   nên tab "Nhật ký TC" luôn hiện mỗi lần mở Excel.
; - Lõi tự cập nhật (tải bản mới, kiểm SHA-256); bộ nạp thay tệp lõi ở lần mở Excel sau.
; - Gỡ cài đặt: xóa giá trị OPEN của add-in và đánh số lại các giá trị sau nó (Excel đọc OPEN, OPEN1, … tới chỗ trống đầu tiên).
;
; Build (scripts/build-vba-release.sh):
;   makensis -DVERSION=x.y.z -DLOADER=<NhatKyThiCong.xlam> -DCORE=<NhatKyThiCong_Core.xlam> -DOUTFILE=<exe> installer/NhatKyTC.nsi

Unicode true
SetCompressor /SOLID lzma
RequestExecutionLevel user

!ifndef VERSION
  !error "Thiếu -DVERSION"
!endif
!ifndef LOADER
  !error "Thiếu -DLOADER"
!endif
!ifndef CORE
  !error "Thiếu -DCORE"
!endif
!ifndef OUTFILE
  !error "Thiếu -DOUTFILE"
!endif

!define APP_NAME "Nhật ký TC (Excel add-in)"
!define APP_DIR "NhatKyTC"
!define PUBLISHER "khoacrab-97"
!define HOMEPAGE "https://github.com/khoacrab-97/Addin_excel_NKTC/releases/latest"
!define LOADER_NAME "NhatKyThiCong.xlam"
!define CORE_NAME "NhatKyThiCong_Core.xlam"
!define OPTIONS_KEY "Software\Microsoft\Office\16.0\Excel\Options"
!define UNINSTALL_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_DIR}"
!define MAX_OPEN 200

Name "${APP_NAME}"
OutFile "${OUTFILE}"
InstallDir "$LOCALAPPDATA\${APP_DIR}"
BrandingText "${APP_NAME} ${VERSION}"
VIProductVersion "${VERSION}.0"
VIAddVersionKey /LANG=0 "ProductName" "${APP_NAME}"
VIAddVersionKey /LANG=0 "ProductVersion" "${VERSION}"
VIAddVersionKey /LANG=0 "FileVersion" "${VERSION}"
VIAddVersionKey /LANG=0 "FileDescription" "Bộ cài ${APP_NAME}"
VIAddVersionKey /LANG=0 "CompanyName" "${PUBLISHER}"
VIAddVersionKey /LANG=0 "LegalCopyright" "${PUBLISHER}"

!include "MUI2.nsh"
!include "StrFunc.nsh"
${StrStr}
${UnStrStr}

!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_WELCOMEPAGE_TITLE "Cài đặt ${APP_NAME}"
!define MUI_WELCOMEPAGE_TEXT "Bộ cài sẽ thêm add-in Nhật ký TC vào Excel trên máy này (chỉ cho tài khoản Windows hiện tại, không cần quyền quản trị).$\r$\n$\r$\nVui lòng LƯU và ĐÓNG Excel trước khi tiếp tục.$\r$\n$\r$\nPhiên bản: ${VERSION}"
!define MUI_FINISHPAGE_TITLE "Đã cài đặt xong"
!define MUI_FINISHPAGE_TEXT "Mở Excel: trên thanh Ribbon có tab Nhật ký TC (luôn hiện mỗi lần mở Excel).$\r$\n$\r$\nAdd-in tự kiểm tra bản mới mỗi ngày (cần Internet); tải xong thì mở lại Excel là dùng bản mới."
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_TEXT "Mở Excel"
!define MUI_FINISHPAGE_RUN_FUNCTION OpenExcel

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Vietnamese"

; Excel đang mở thì giữ tệp add-in và ghi đè khóa OPEN khi thoát -> yêu cầu đóng Excel (lớp cửa sổ chính: XLMAIN).
!macro RequireExcelClosed PREFIX
  ${PREFIX}_check:
    FindWindow $0 "XLMAIN"
    StrCmp $0 0 ${PREFIX}_closed
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Excel đang mở. Vui lòng lưu và đóng tất cả cửa sổ Excel rồi bấm Retry." IDRETRY ${PREFIX}_check
    Abort
  ${PREFIX}_closed:
!macroend

; Tên giá trị thứ $R0 (0 -> OPEN, 1 -> OPEN1, …) vào $R1.
!macro OpenValueName
  StrCmp $R0 0 0 +3
    StrCpy $R1 "OPEN"
    Goto +2
  StrCpy $R1 "OPEN$R0"
!macroend

Function OpenExcel
  ExecShell "open" "excel.exe"
FunctionEnd

Function .onInit
  !insertmacro RequireExcelClosed install
FunctionEnd

Function un.onInit
  !insertmacro RequireExcelClosed uninstall
FunctionEnd

; Ghi bộ nạp vào giá trị OPEN đang trỏ tới add-in này (cài lại / cập nhật), hoặc giá trị trống đầu tiên.
Function RegisterAddin
  StrCpy $R0 0
  register_loop:
    !insertmacro OpenValueName
    ClearErrors
    ReadRegStr $R3 HKCU "${OPTIONS_KEY}" "$R1"
    IfErrors register_write
    ${StrStr} $R4 $R3 "${LOADER_NAME}"
    StrCmp $R4 "" 0 register_write
    IntOp $R0 $R0 + 1
    IntCmp $R0 ${MAX_OPEN} register_full register_loop register_full
  register_write:
    WriteRegStr HKCU "${OPTIONS_KEY}" "$R1" '/R "$INSTDIR\${LOADER_NAME}"'
    Return
  register_full:
    MessageBox MB_ICONEXCLAMATION "Không đăng ký được add-in với Excel (quá nhiều add-in tự mở). Hãy thêm thủ công: Excel → File → Options → Add-ins → Go… → Browse… → $INSTDIR\${LOADER_NAME}."
FunctionEnd

; Xóa giá trị OPEN của add-in này và dồn các giá trị phía sau lên một bậc.
Function un.UnregisterAddin
  StrCpy $R0 0
  find_loop:
    !insertmacro OpenValueName
    ClearErrors
    ReadRegStr $R3 HKCU "${OPTIONS_KEY}" "$R1"
    IfErrors find_done
    ${UnStrStr} $R4 $R3 "${LOADER_NAME}"
    StrCmp $R4 "" 0 shift_loop
    IntOp $R0 $R0 + 1
    IntCmp $R0 ${MAX_OPEN} find_done find_loop find_done
  shift_loop:
    ; $R1 = tên giá trị hiện tại (vị trí $R0); đọc giá trị kế tiếp vào $R3.
    StrCpy $R5 $R1
    IntOp $R0 $R0 + 1
    !insertmacro OpenValueName
    ClearErrors
    ReadRegStr $R3 HKCU "${OPTIONS_KEY}" "$R1"
    IfErrors shift_last
    WriteRegStr HKCU "${OPTIONS_KEY}" "$R5" "$R3"
    IntCmp $R0 ${MAX_OPEN} shift_last shift_loop shift_last
  shift_last:
    DeleteRegValue HKCU "${OPTIONS_KEY}" "$R5"
  find_done:
FunctionEnd

Section "Install"
  SetOutPath "$INSTDIR"
  ; Bản cập nhật tải dở của lần cài trước không còn dùng.
  Delete "$INSTDIR\NhatKyThiCong_Core.pending"
  Delete "$INSTDIR\NhatKyThiCong_Core.pending.ok"
  Delete "$INSTDIR\NhatKyThiCong_Core.bak"
  File "/oname=${LOADER_NAME}" "${LOADER}"
  File "/oname=${CORE_NAME}" "${CORE}"
  File "/oname=icon.ico" "icon.ico"
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  Call RegisterAddin

  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "Publisher" "${PUBLISHER}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "URLInfoAbout" "${HOMEPAGE}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayIcon" "$INSTDIR\icon.ico"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "${UNINSTALL_KEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoRepair" 1
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "EstimatedSize" 400
SectionEnd

Section "Uninstall"
  Call un.UnregisterAddin
  DeleteRegKey HKCU "${UNINSTALL_KEY}"
  Delete "$INSTDIR\${LOADER_NAME}"
  Delete "$INSTDIR\${CORE_NAME}"
  Delete "$INSTDIR\NhatKyThiCong_Core.pending"
  Delete "$INSTDIR\NhatKyThiCong_Core.pending.ok"
  Delete "$INSTDIR\NhatKyThiCong_Core.bak"
  Delete "$INSTDIR\settings.txt"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
SectionEnd
