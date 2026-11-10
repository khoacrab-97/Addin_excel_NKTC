# Nhật ký TC — bản add-in VBA (`.xlam`)

Bản chuyển đổi từ add-in Office.js sang add-in VBA, cùng kiểu với TienBangChu (tab luôn hiện khi mở Excel, ô chọn vùng dùng hộp chọn vùng gốc của Excel, không cần Internet khi dùng). Tên tab: **Nhật ký TC**. Phát hành chính thức 3.0.0 (10/11/2026, ROADMAP Phase 12); bản Office.js đã gỡ khỏi repo (Phase 13).

## Cấu trúc

| Thư mục / tệp | Nội dung |
|---|---|
| `src/*.bas`, `src/*.cls` | Mã nguồn VBA, viết bằng UTF-8 (tiếng Việt có dấu trong chuỗi và chú thích) |
| `src/*.frmcode` | Mã của UserForm (sự kiện nút, kiểm tra nhập) |
| `src/ThisWorkbook.doccode`, `loader/ThisWorkbook.doccode` | Mã ThisWorkbook (Workbook_Open) của lõi / bộ nạp |
| `forms/*.json` | Bố cục UserForm: tiêu đề, kích thước, các ô (Label, TextBox, ComboBox, OptionButton, CommandButton…) và thuộc tính |
| `tools/vba_ascii.py` | Đổi mã nguồn sang ASCII trước khi nạp vào trình soạn VBA: chuỗi tiếng Việt → `U("\uXXXX")`, chú thích bỏ dấu |
| `tools/test_vba_ascii.py` | Kiểm thử công cụ trên (`python3 -m unittest test_vba_ascii`, chạy trong `vba/tools`) |
| `loader/` | Bộ nạp `NhatKyThiCong.xlam` (thay lõi khi có bản cập nhật đã tải, rồi mở lõi) |
| `build/src-ascii/`, `build/loader-ascii/` | Mã ASCII sinh từ `src/`, `loader/` (không sửa tay; CI kiểm tra luôn khớp) |
| `release/` | 3 tệp của bản phát hành (tải lên từ Windows) — xem `release/README.md` |
| `ribbon/customUI14.xml` | Tab "Nhật ký TC"; ảnh nút lấy từ `assets/ribbon/*-32.png` |
| `build/build.cmd`, `build/build.ps1` | Build trên Windows có Excel: nạp mã, tạo UserForm từ `forms/*.json` + `.frmcode`, chép sheet "Mẫu NKTC" từ `templates/Mau_NKTC.xlsx` vào add-in, biên dịch, lưu `dist/NhatKyThiCong_Core.xlam`, gắn Ribbon, chạy kiểm thử; kết quả ở `dist/ket-qua-build.txt` |

## Build (Windows, có Excel)

1. Một lần: Excel → **File → Options → Trust Center → Trust Center Settings → Macro Settings** → tích **Trust access to the VBA project object model**.
2. Sau khi sửa `src/` hoặc `loader/`: chạy `python3 vba/tools/vba_ascii.py vba/src vba/build/src-ascii` và `python3 vba/tools/vba_ascii.py vba/loader vba/build/loader-ascii` (CI báo lỗi nếu quên).
3. Nháy đúp `vba/build/build.cmd`. Lỗi biên dịch → Excel hiện hộp thoại báo dòng lỗi. Kết quả trong `vba/dist/`: `NhatKyThiCong.xlam` (bộ nạp), `NhatKyThiCong_Core.xlam` (lõi), `version.txt`, `ket-qua-build.txt`.

## Phát hành

1. Tăng `APP_VERSION` trong `src/modNKTC_Ribbon.bas` (ví dụ 3.0.1), sinh lại mã ASCII, build như trên (phải **ĐẠT**).
2. Tải 3 tệp `vba/dist/NhatKyThiCong.xlam`, `NhatKyThiCong_Core.xlam`, `version.txt` lên `vba/release/` (GitHub → Add file → Upload files).
3. Đưa lên `main`: CI kiểm tra SHA-256, đăng kênh cập nhật (GitHub Pages `/vba/`), dựng bộ cài `NhatKyTC_Setup_<phiên bản>.exe`, tạo Release `vba-v<phiên bản>` (tiền tố riêng, không trùng tag của bản Office.js); CI dừng nếu `version.txt` khác `APP_VERSION`. Máy đã cài tự tải bản mới trong vòng một ngày.

## Cài đặt và cập nhật trên máy người dùng

- Bộ cài (`installer/NhatKyTC.nsi`, không cần quyền quản trị) chép bộ nạp + lõi vào `%LOCALAPPDATA%\NhatKyTC` và ghi giá trị `OPEN…` trong `HKCU\Software\Microsoft\Office\16.0\Excel\Options` → Excel mở bộ nạp mỗi lần khởi động. Gỡ cài đặt: xóa giá trị đó và dồn các giá trị phía sau.
- Lõi kiểm tra `version.txt` mỗi ngày một lần (30 giây sau khi mở Excel) và khi bấm **Tiện ích → Kiểm tra cập nhật**; bản mới phải khớp kích thước, chữ ký ZIP, SHA-256 rồi mới được lưu thành `.pending` + `.pending.ok`. Lần mở Excel sau, bộ nạp kiểm lại SHA-256 và thay tệp lõi (giữ bản cũ ở `.bak`).
- `settings.txt` cạnh add-in: `LastCheck=` (lần kiểm tra gần nhất), `UpdateUrl=` (kênh thay thế để kiểm thử; chỉ nhận https hoặc 127.0.0.1/localhost).

## Module

| Module | Nội dung | Tương ứng bản Office.js cũ (lịch sử Git, tag `v2.0.1`) |
|---|---|---|
| `modNKTC_Text` | `U()`, bỏ dấu, chuẩn hóa tiêu đề | `src/core/text` |
| `modNKTC_Date` | Ngày (số nguyên hệ 1900), đọc ngày dạng chữ, mẫu "Ngày … tháng … năm …" | `src/core/dates` |
| `modNKTC_DailyLog` | F1: nhận diện cột, đọc công việc, trải theo ngày, tên sheet | `src/core/dailyLog` |
| `modNKTC_Mapping` | F2: gán cột, danh sách ngày, chuyển ngày, khớp cột theo tên, giá trị ghi ô | `src/core/mapping` |
| `modNKTC_Config` | F2: địa chỉ ô đích, dòng gán, cấu hình gán (lưu/đọc dạng chữ), khớp lại cột theo tên | `src/core/address/reference`, `src/core/mapping/{formRows,config,headerMatch}` |
| `modNKTC_Excel` | Lớp Excel: đọc vùng (`Value2`), tạo sheet NK_TheoNgay, tạo bảng nhập tiến độ (không ghi đè) | `src/excel/*` |
| `modNKTC_Commands` | Lệnh Ribbon: Tạo bảng (F4), Xuất theo ngày (F1) | `src/commands/*` |
| `modNKTC_FormCommands` | Lệnh Ribbon F2: Chèn form mẫu, Xuất vào form, Chuyển ngày; cấu hình lưu trong file (CustomXMLPart `urn:nhatkytc:config:1`); Hướng dẫn | `src/commands/{formCommands,templateCommands}`, `src/excel/settingsStore` |
| `modNKTC_UI` | Hộp thông báo tiếng Việt (`ShowInfo`, `ShowMessage` qua `MessageBoxW` của Windows) | `src/commands/messages.ts` |
| `modNKTC_Update` | Tự cập nhật: đọc version.txt, tải, kiểm tra kích thước / ZIP / SHA-256 (CryptoAPI), lưu `.pending` | (bản Office.js tự cập nhật qua GitHub Pages) |
| `modNKTC_Loader` (bộ nạp) | Thay lõi bằng bản đã tải (kiểm lại SHA-256), mở lõi | — |
| `modNKTC_Ribbon` | Nhận lệnh từ nút Ribbon (`RB_Action`), Phiên bản | `manifest.xml` + `commands.ts` |
| `frmTaoBang`, `frmXuatNgay`, `frmXuatForm` | Hộp thoại Tạo bảng, Xuất nhật ký theo ngày, Xuất vào form (ô chọn vùng: ô nhập + nút **Chọn**) | `src/dialog/views/*` |
| `modNKTC_Tests` | Kiểm thử phần lõi: `NKTC_RunTests()` → `PASS n` hoặc danh sách lỗi | `src/core/**/*.test.ts` |

## Quy ước

- Phần lõi (`modNKTC_Text`, `modNKTC_Date`, `modNKTC_DailyLog`, `modNKTC_Mapping`) không gọi đối tượng Excel; đọc dữ liệu ô dạng `Range.Value2`.
- `Const` không chứa chuỗi tiếng Việt (VBA không cho gọi hàm trong `Const`) — công cụ chuyển đổi báo lỗi nếu vi phạm.
- Không dùng `MsgBox`/`InputBox` của VBA (đổi sang ANSI, mất dấu tiếng Việt): dùng `ShowInfo`/`ShowMessage` và `Application.InputBox`. Kiểm thử `test_vba_ascii.py` chặn vi phạm.
- Không dùng ActiveX trên UserForm (kể cả RefEdit): Microsoft 365 mới chặn chèn ActiveX. Chọn vùng bằng `Application.InputBox(Type:=8)` qua `PickRange`.
- Không dùng `VBScript.RegExp` (Microsoft đang loại bỏ VBScript khỏi Windows).
