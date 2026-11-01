# Hướng dẫn cài đặt Add-in "Nhật ký thi công"

> Áp dụng: Excel Desktop trên Windows — Microsoft 365 hoặc Excel 2016 trở về sau.
> Link gửi người dùng: <https://khoacrab-97.github.io/Addin_excel_NKTC/>
> Tải trực tiếp bộ cài: <https://khoacrab-97.github.io/Addin_excel_NKTC/NhatKyThiCong_Setup.exe>
> GitHub Releases (bộ cài theo từng phiên bản, `NhatKyThiCong_Setup_<phiên bản>.exe`): <https://github.com/khoacrab-97/Addin_excel_NKTC/releases/latest>

## 1. Điều kiện

| Yêu cầu | Ghi chú |
|---|---|
| Excel Desktop Windows (Microsoft 365 / 2016 / 2019 / 2021 / 2024) | Excel 2016 bản mua một lần chạy được nhưng không tự đặt độ rộng cột (xem mục 6). |
| Kết nối Internet khi mở add-in | Add-in tải giao diện từ `khoacrab-97.github.io` và thư viện Office từ `appsforoffice.microsoft.com`. |
| Tài khoản Windows thường | Bộ cài **không cần quyền Administrator**; chỉ cài cho tài khoản Windows đang đăng nhập. |

## 2. Cài đặt bằng bộ cài (khuyến nghị)

1. **Đóng hẳn Excel** (bộ cài sẽ nhắc nếu Excel còn mở).
2. Mở <https://khoacrab-97.github.io/Addin_excel_NKTC/> → **Tải bộ cài (Windows)** → chạy `NhatKyThiCong_Setup.exe`.
3. Nếu Windows hiện **"Windows protected your PC"** (SmartScreen): bấm **More info** → **Run anyway**. Cảnh báo này xuất hiện vì bộ cài chưa có chữ ký số thương mại.
4. Bấm **Cài đặt** → **Hoàn tất**.
5. Mở Excel → tab **Home** xuất hiện nhóm **Nhật ký thi công** → bấm nút để mở khung làm việc.
6. Kiểm tra: tab **Tiện ích** hiển thị phiên bản add-in, phiên bản Office và mức **ExcelApi tối đa**.

Nếu không thấy nút trên tab Home: **Insert** (Chèn) → **My Add-ins** → thẻ **Developer Add-ins** (hoặc **MY ADD-INS**) → chọn **Nhật ký thi công** → **Add**.

Bộ cài thực hiện (chỉ trong phạm vi tài khoản người dùng):

| Hạng mục | Vị trí |
|---|---|
| Chép `manifest.xml`, icon, trình gỡ cài đặt | `%LOCALAPPDATA%\NhatKyThiCong\` |
| Đăng ký add-in với Office | `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` (tên giá trị = ID add-in, dữ liệu = đường dẫn `manifest.xml`) |
| Mục gỡ cài đặt trong Windows | `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NhatKyThiCong` |

## 3. Cập nhật

| Loại cập nhật | Cách nhận |
|---|---|
| Sửa lỗi / thêm chức năng (thường gặp) | **Tự động**: mỗi lần mở khung add-in, Excel tải bản mới nhất từ máy chủ phát hành. Khi có bản mới, add-in hiện thông báo "Đã có phiên bản mới" đầu khung; tab **Tiện ích** → **Kiểm tra cập nhật** để kiểm tra bất cứ lúc nào. |
| Có thay đổi manifest (thông báo ghi rõ "cần cài lại") | Đóng Excel → tải và chạy lại `NhatKyThiCong_Setup.exe` (cài đè, không cần gỡ trước) → mở lại Excel. |

## 4. Gỡ cài đặt

**Settings** (Cài đặt) → **Apps** → **Installed apps** (Ứng dụng đã cài) → **Nhật ký thi công (Excel add-in)** → **Uninstall**. Đóng Excel trước khi gỡ.

## 5. Cách cài thay thế: thư mục chia sẻ (Shared Folder Catalog)

Dùng khi máy không chạy được bộ cài (chính sách công ty chặn file `.exe`) hoặc Excel không nhận add-in sau khi cài.

1. Tạo thư mục, ví dụ `C:\NKTC_Addin` → chuột phải → **Properties** → thẻ **Sharing** → **Share…** → chọn tài khoản của mình → **Share**. Ghi lại **đường dẫn mạng**, ví dụ `\\DESKTOP-ABC123\NKTC_Addin` (Excel chỉ nhận đường dẫn mạng, không nhận `C:\...`).
2. Mở trang phát hành → mục **Cài cho cả nhóm qua thư mục chia sẻ (nâng cao)** → tải **manifest.xml** → lưu vào thư mục trên.
3. Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings…** → **Trusted Add-in Catalogs** → dán đường dẫn mạng vào **Catalog Url** → **Add catalog** → tích **Show in Menu** → **OK** → **OK**.
4. Đóng hẳn Excel và mở lại → **Insert** → **My Add-ins** → thẻ **SHARED FOLDER** → **Nhật ký thi công** → **Add**.

> Không dùng đồng thời hai cách cài trên cùng một máy — hai bản có cùng ID add-in.

## 6. Khác biệt theo phiên bản Excel

| Chức năng | Microsoft 365 / 2019+ | Excel 2016 (mua một lần) |
|---|---|---|
| F1 – Xuất nhật ký theo ngày | Đầy đủ, tự đặt độ rộng cột | Đầy đủ; độ rộng cột của sheet kết quả chỉnh tay |
| F2 – Xuất theo form mẫu | Đầy đủ | Đầy đủ |

## 7. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Bộ cài báo "Excel đang mở" | Đóng mọi cửa sổ Excel (kiểm tra cả biểu tượng Excel trên thanh tác vụ) → **Retry**. |
| Cài xong không thấy nút trên tab Home | Đóng hẳn Excel và mở lại; nếu vẫn không thấy → **Insert → My Add-ins → Developer Add-ins** (mục 2). Vẫn không có → dùng cách cài mục 5 và báo lại cho người phát triển. |
| Không thấy thẻ **SHARED FOLDER** (cách cài mục 5) | Đã tích **Show in Menu** và đã khởi động lại Excel; đường dẫn phải là đường dẫn mạng `\\...`. |
| Khung add-in trắng / báo lỗi tải | Kiểm tra Internet; thử mở <https://khoacrab-97.github.io/Addin_excel_NKTC/taskpane.html> bằng trình duyệt. |
| Đã cập nhật nhưng vẫn thấy bản cũ | Đóng Excel → xóa toàn bộ nội dung thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` → mở lại Excel. |
| Nút trên giao diện không có biểu tượng | Mạng chặn CDN biểu tượng của Microsoft; chức năng vẫn dùng bình thường (các nút điều hướng ngày dùng ký tự « ‹ › »). |

## 8. Dành cho người phát triển

- Chạy bản phát triển: `npm start` (sideload manifest `localhost`). **Không** dùng đồng thời với bản đã cài bằng bộ cài — hai bản có cùng ID add-in; gỡ bản cài trước (mục 4).
- Bộ cài: `installer/NhatKyThiCong.nsi` (NSIS 3). Build: `npm run build` rồi `npm run build:installer` → `dist/NhatKyThiCong_Setup.exe` (cần `makensis`). CI và workflow phát hành tự build.
- Phát hành: workflow `.github/workflows/deploy-pages.yml` chạy khi merge vào `main` (hoặc chạy tay trong tab **Actions**): đưa add-in + bộ cài lên GitHub Pages, rồi tạo GitHub Release `v<phiên bản>` kèm `NhatKyThiCong_Setup_<phiên bản>.exe` (bỏ qua nếu phiên bản đã có Release — muốn phát hành Release mới phải tăng `version` trong `package.json`). Yêu cầu một lần: **Settings → Pages → Source: GitHub Actions**.
