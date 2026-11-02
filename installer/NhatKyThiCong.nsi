; Bộ cài add-in "Nhật ký thi công" cho Excel Desktop Windows.
; - Không cần quyền quản trị: chỉ ghi vào %LOCALAPPDATA% và registry HKCU.
; - Đăng ký manifest với Excel theo cơ chế sideload của Microsoft (office-addin-dev-settings):
;   HKCU\Software\Microsoft\Office\16.0\Wef\Developer, tên giá trị = ID add-in, dữ liệu = đường dẫn manifest.
; - Kích hoạt: mở workbook NhatKyThiCong_BatDau.xlsx (có tham chiếu tới add-in, sinh bằng generateSideloadFile
;   của Microsoft) → Excel nạp add-in, mở khung làm việc và hiện nút trên tab Home. Có lối tắt trong Start menu.
; - Mã add-in chạy từ GitHub Pages nên tự cập nhật; chạy lại bộ cài chỉ cần khi manifest thay đổi.
;
; Build (scripts/build-installer.js):
;   makensis -DVERSION=x.y.z -DMANIFEST=<dist/manifest.xml> -DSTARTER=<workbook kích hoạt> -DOUTFILE=<exe> installer/NhatKyThiCong.nsi

Unicode true
SetCompressor /SOLID lzma
RequestExecutionLevel user

!ifndef VERSION
  !error "Thiếu -DVERSION"
!endif
!ifndef MANIFEST
  !error "Thiếu -DMANIFEST"
!endif
!ifndef STARTER
  !error "Thiếu -DSTARTER"
!endif
!ifndef OUTFILE
  !error "Thiếu -DOUTFILE"
!endif

!define APP_NAME "Nhật ký thi công (Excel add-in)"
!define APP_ID "2b55386a-bc23-4f8d-81e4-137b8271b9ea"
!define APP_DIR "NhatKyThiCong"
!define PUBLISHER "khoacrab-97"
!define HOMEPAGE "https://khoacrab-97.github.io/Addin_excel_NKTC/"
!define WEF_KEY "Software\Microsoft\Office\16.0\Wef\Developer"
!define STARTER_NAME "NhatKyThiCong_BatDau.xlsx"
!define SHORTCUT "$SMPROGRAMS\Nhật ký thi công (Excel).lnk"
!define UNINSTALL_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_DIR}"

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
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_WELCOMEPAGE_TITLE "Cài đặt ${APP_NAME}"
!define MUI_WELCOMEPAGE_TEXT "Bộ cài sẽ thêm add-in Nhật ký thi công vào Excel trên máy này (chỉ cho tài khoản Windows hiện tại, không cần quyền quản trị).$\r$\n$\r$\nVui lòng LƯU và ĐÓNG Excel trước khi tiếp tục.$\r$\n$\r$\nPhiên bản: ${VERSION}"
!define MUI_FINISHPAGE_TITLE "Đã cài đặt xong"
!define MUI_FINISHPAGE_TEXT "Bấm Hoàn tất để mở Excel và kích hoạt add-in (lần đầu Excel có thể hỏi tin cậy add-in → chọn Trust/Tin cậy).$\r$\n$\r$\nSau đó dùng nút Nhật ký thi công trên tab Home. Mở lại lần sau: Start menu → Nhật ký thi công (Excel).$\r$\n$\r$\nAdd-in tự cập nhật mỗi lần mở; cần kết nối Internet."
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_TEXT "Mở Excel và kích hoạt add-in"
!define MUI_FINISHPAGE_RUN_FUNCTION OpenStarterWorkbook
!define MUI_FINISHPAGE_LINK "Hướng dẫn và trang phát hành"
!define MUI_FINISHPAGE_LINK_LOCATION "${HOMEPAGE}"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Vietnamese"

; Excel đang mở thì không nạp manifest mới → yêu cầu đóng Excel (lớp cửa sổ chính của Excel: XLMAIN).
!macro RequireExcelClosed PREFIX
  ${PREFIX}_check:
    FindWindow $0 "XLMAIN"
    StrCmp $0 0 ${PREFIX}_closed
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Excel đang mở. Vui lòng lưu và đóng tất cả cửa sổ Excel rồi bấm Retry." IDRETRY ${PREFIX}_check
    Abort
  ${PREFIX}_closed:
!macroend

Function OpenStarterWorkbook
  ExecShell "open" "$INSTDIR\${STARTER_NAME}"
FunctionEnd

Function .onInit
  !insertmacro RequireExcelClosed install
FunctionEnd

Function un.onInit
  !insertmacro RequireExcelClosed uninstall
FunctionEnd

Section "Install"
  SetOutPath "$INSTDIR"
  File "/oname=manifest.xml" "${MANIFEST}"
  File "/oname=icon.ico" "icon.ico"
  File "/oname=${STARTER_NAME}" "${STARTER}"
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  WriteRegStr HKCU "${WEF_KEY}" "${APP_ID}" "$INSTDIR\manifest.xml"

  SetShellVarContext current
  CreateShortCut "${SHORTCUT}" "$INSTDIR\${STARTER_NAME}" "" "$INSTDIR\icon.ico"

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
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "EstimatedSize" 200
SectionEnd

Section "Uninstall"
  DeleteRegValue HKCU "${WEF_KEY}" "${APP_ID}"
  DeleteRegKey HKCU "${UNINSTALL_KEY}"
  SetShellVarContext current
  Delete "${SHORTCUT}"
  Delete "$INSTDIR\${STARTER_NAME}"
  Delete "$INSTDIR\manifest.xml"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
SectionEnd
