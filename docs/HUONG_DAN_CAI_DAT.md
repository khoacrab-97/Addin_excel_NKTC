# Hướng dẫn cài đặt Add-in "Nhật ký TC" (từ 3.0.0)

> Áp dụng: Excel Desktop trên Windows — Excel 2016, 2019, 2021, 2024 hoặc Microsoft 365 (32-bit và 64-bit).
> Tải bộ cài: [**Releases mới nhất**](https://github.com/khoacrab-97/Addin_excel_NKTC/releases/latest) → `NhatKyTC_Setup_<phiên bản>.exe`.
> Bản cũ "Nhật ký thi công" (Office.js, 1.x – 2.x) đã ngừng hỗ trợ — xem mục 7 để gỡ.

## 1. Điều kiện

| Yêu cầu | Ghi chú |
|---|---|
| Excel Desktop Windows (2016 trở về sau, hoặc Microsoft 365) | Không chạy trên Excel Online, Mac, điện thoại. |
| Tài khoản Windows thường | Bộ cài **không cần quyền Administrator**; chỉ cài cho tài khoản Windows đang đăng nhập. |
| Internet | **Không cần** khi dùng. Chỉ cần khi add-in tự kiểm tra bản mới (mỗi ngày một lần) hoặc khi bấm **Kiểm tra cập nhật**. |

## 2. Cài đặt

1. **Lưu và đóng hẳn Excel** (bộ cài sẽ nhắc nếu Excel còn mở).
2. Tải `NhatKyTC_Setup_<phiên bản>.exe` từ Releases và chạy.
3. Nếu Windows hiện **"Windows protected your PC"** (SmartScreen): bấm **More info** → **Run anyway**. Cảnh báo này xuất hiện vì bộ cài chưa có chữ ký số thương mại. Muốn chắc chắn tệp tải về không bị thay đổi: mở PowerShell tại thư mục chứa tệp, chạy
   `Get-FileHash .\NhatKyTC_Setup_<phiên bản>.exe -Algorithm SHA256` và so với mã SHA-256 ghi trong Release.
4. Bấm **Next** đến hết. Ở trang cuối có thể giữ dấu tích **Mở Excel**.
5. Mở Excel: thanh Ribbon có tab **Nhật ký TC** — tab luôn hiện mỗi lần mở Excel, với mọi file.
6. Kiểm tra: **Nhật ký TC → Tiện ích ▾ → Phiên bản** hiển thị phiên bản add-in, phiên bản Excel, thư mục cài đặt và kết quả tự kiểm tra phần lõi (`PASS …`).

Tab **Nhật ký TC**:

| Nhóm | Nút |
|---|---|
| Tiến độ | Tạo bảng · Xuất theo ngày |
| Form mẫu | Chèn form mẫu · Xuất vào form |
| Chuyển ngày | Ngày đầu · Ngày trước · Ngày sau · Ngày cuối |
| In | In nhật ký |
| Tiện ích | Menu **Tiện ích ▾**: Hướng dẫn · Phiên bản · Kiểm tra cập nhật |

Bộ cài thực hiện (chỉ trong phạm vi tài khoản người dùng):

| Hạng mục | Vị trí |
|---|---|
| Bộ nạp `NhatKyThiCong.xlam`, lõi `NhatKyThiCong_Core.xlam` | `%LOCALAPPDATA%\NhatKyTC\` |
| Đăng ký add-in tự mở khi khởi động Excel | `HKCU\Software\Microsoft\Office\16.0\Excel\Options` — giá trị `OPEN`, `OPEN1`, … trỏ tới bộ nạp |
| Mục gỡ cài đặt | `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NhatKyTC` |

## 3. Cập nhật

- Add-in tự kiểm tra bản mới **mỗi ngày một lần**, khoảng 30 giây sau khi mở Excel (im lặng nếu không có gì mới). Muốn kiểm tra ngay: **Tiện ích ▾ → Kiểm tra cập nhật**.
- Có bản mới: add-in tải về, kiểm tra kích thước, định dạng tệp và mã **SHA-256**; đạt mới giữ lại và báo *"Đã tải bản cập nhật … Hãy đóng và mở lại Excel"*. Lần mở Excel sau, bộ nạp kiểm tra lại SHA-256 rồi thay bản cũ (bản cũ giữ ở `NhatKyThiCong_Core.bak`).
- Tệp tải về hỏng hoặc không khớp mã → bị bỏ, add-in tiếp tục dùng bản đang có.
- Không cần chạy lại bộ cài khi cập nhật, trừ khi Release ghi rõ "cần cài lại".

## 4. Gỡ cài đặt

Đóng Excel → **Settings → Apps → Installed apps** → **Nhật ký TC (Excel add-in)** → **Uninstall**. Bộ gỡ xóa thư mục `%LOCALAPPDATA%\NhatKyTC` và giá trị `OPEN…` của add-in (các add-in khác giữ nguyên). Dữ liệu và cấu hình gán trong các file Excel của bạn không bị ảnh hưởng.

## 5. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Không thấy tab **Nhật ký TC** | Excel → **File → Options → Add-ins** → ô *Manage* chọn **Excel Add-ins** → **Go…** → tích **NhatKyThiCong** (nếu không có: **Browse…** → `%LOCALAPPDATA%\NhatKyTC\NhatKyThiCong.xlam`). Nếu nằm ở **Disabled Items** (ô *Manage* → *Disabled Items* → **Go…**) thì chọn và **Enable**. |
| Thông báo *"Không tìm thấy tệp NhatKyThiCong_Core.xlam"* | Chạy lại bộ cài. |
| **Kiểm tra cập nhật** báo lỗi mạng / HTTP | Kiểm tra Internet; add-in vẫn dùng bình thường với bản hiện có. |
| **Kiểm tra cập nhật** báo *HTTP 404 … /vba/version.txt* | Lỗi phía máy chủ phát hành (không phải máy người dùng) — báo người bảo trì xử lý theo mục 8. Add-in vẫn dùng bình thường với bản hiện có. |
| Còn tab "Nhật ký thi công" (bản cũ) | Gỡ bản cũ theo mục 7. |

## 6. Ghi chú bảo mật

- Add-in chỉ đọc vùng bạn chọn và chỉ ghi vào: sheet mới do add-in tạo, vùng bảng trống bạn chọn, và các ô đích bạn đã gán trong **Xuất vào form** (ô có công thức/dữ liệu sẽ được hỏi trước khi ghi đè).
- Cấu hình gán lưu ngay trong file Excel (không gửi đi đâu). Add-in không thu thập dữ liệu; kết nối mạng duy nhất là tải `version.txt` và bản cập nhật từ `https://khoacrab-97.github.io/Addin_excel_NKTC/vba/`.

## 7. Gỡ bản cũ "Nhật ký thi công" (Office.js, 1.x – 2.x)

Bản cũ tải mã từ Internet mỗi lần mở và đã ngừng hỗ trợ từ 10/11/2026 (các nút không còn chạy). Gỡ như sau:

1. Đóng Excel → **Settings → Apps → Installed apps** → **Nhật ký thi công (Excel add-in)** → **Uninstall** (lưu ý: không gỡ nhầm **Nhật ký TC**).
2. Nếu còn thấy tab cũ: Excel → **Insert → My Add-ins** (Phần bổ trợ của tôi) → thẻ **Developer Add-ins** → chuột phải **Nhật ký thi công** → **Remove**.
3. Cấu hình gán của bản cũ lưu theo cách khác, nên với file nhật ký cũ hãy bấm **Xuất vào form** để gán lại một lần.

## 8. Dành cho người bảo trì: kênh phát hành (GitHub Pages)

Kênh tự cập nhật `https://khoacrab-97.github.io/Addin_excel_NKTC/vba/` (`version.txt` + `NhatKyThiCong_Core.xlam`), trang tải về và GitHub Release `vba-v<phiên bản>` đều do workflow **Deploy GitHub Pages** (`.github/workflows/deploy-pages.yml`) tạo mỗi khi có thay đổi vào `main`.

**Cài đặt bắt buộc:** repo → **Settings → Pages → Build and deployment → Source = "GitHub Actions"**.

- **Không** để *"Deploy from a branch"*: khi đó mỗi lần merge, GitHub chạy thêm quy trình *pages build and deployment* đăng nguyên thư mục repo (không có `vba/version.txt`). Quy trình nào xong sau sẽ ghi đè quy trình kia → kênh cập nhật có lúc trả **HTTP 404**, tất cả máy người dùng báo *"Không kiểm tra được bản cập nhật"* (sự cố ngày 11/11/2026 khi phát hành 3.1.0).
- Kiểm tra sau mỗi lần phát hành: mở `https://khoacrab-97.github.io/Addin_excel_NKTC/vba/version.txt` (Ctrl+F5) → dòng đầu phải là `version=<phiên bản mới>`; trang **Settings → Pages** ghi *"last deployed … by the Deploy GitHub Pages workflow"*.
- Nếu `version.txt` báo 404: kiểm tra lại Source như trên, rồi **Actions → Deploy GitHub Pages → Run workflow** (nhánh `main`) để đăng lại. Chạy lại không tạo Release trùng (phiên bản đã có Release sẽ được bỏ qua).

