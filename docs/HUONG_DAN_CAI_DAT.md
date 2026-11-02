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
4. Bấm **Cài đặt**. Ở trang cuối giữ dấu tích **Mở Excel và kích hoạt add-in** → **Hoàn tất**.
5. Excel mở workbook `NhatKyThiCong_BatDau.xlsx` kèm khung chào mừng (chỉ hiện lần này, đóng được). Lần đầu Excel có thể hỏi tin cậy add-in → chọn **Trust** / **Tin cậy**. Từ đó thanh Ribbon có tab **Nhật ký thi công**; mở bất kỳ file nào cũng dùng được.
6. Kiểm tra: tab **Nhật ký thi công** → nút **Phiên bản** hiển thị phiên bản add-in, phiên bản Office và mức **ExcelApi tối đa**.

Tab **Nhật ký thi công** (từ 1.3.0, không còn khung bên phải):

| Nhóm | Nút |
|---|---|
| Xuất nhật ký | Xuất theo ngày |
| Form mẫu | Vùng nguồn · Gán cột · Danh sách gán |
| Chuyển ngày | Ngày đầu · Ngày trước · Ngày sau · Ngày cuối · Đến ngày… |
| Tiện ích | Hướng dẫn · Phiên bản |

Nếu không thấy tab **Nhật ký thi công**: **Start menu** → **Nhật ký thi công (Excel)** (mở lại workbook kích hoạt), hoặc **Insert** (Chèn) → **My Add-ins** → thẻ **Developer Add-ins** (hoặc **MY ADD-INS**) → chọn **Nhật ký thi công** → **Add**.

> Vì sao cần bước kích hoạt: bộ cài đăng ký add-in trong registry; Excel chỉ nạp add-in đã đăng ký khi mở một workbook có tham chiếu tới add-in (cách `npm start` của Microsoft cũng dùng) hoặc khi người dùng thêm từ **My Add-ins**. Bản 1.2.0 thiếu bước này.

Bộ cài thực hiện (chỉ trong phạm vi tài khoản người dùng):

| Hạng mục | Vị trí |
|---|---|
| Chép `manifest.xml`, icon, workbook kích hoạt `NhatKyThiCong_BatDau.xlsx`, trình gỡ cài đặt | `%LOCALAPPDATA%\NhatKyThiCong\` |
| Lối tắt **Nhật ký thi công (Excel)** (mở workbook kích hoạt) | Start menu của người dùng |
| Đăng ký add-in với Office | `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` (tên giá trị = ID add-in, dữ liệu = đường dẫn `manifest.xml`) |
| Mục gỡ cài đặt trong Windows | `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NhatKyThiCong` |

## 3. Cập nhật

| Loại cập nhật | Cách nhận |
|---|---|
| Sửa lỗi / thêm chức năng (thường gặp) | **Tự động**: mỗi lần Excel nạp add-in, mã add-in được tải bản mới nhất từ máy chủ phát hành. Nút **Phiên bản** kiểm tra và hướng dẫn nếu còn chạy bản cũ. |
| Có thay đổi manifest (ghi rõ "cần cài lại" trong Release; ví dụ **1.3.0** chuyển sang tab Ribbon) | Đóng Excel → tải và chạy lại bộ cài mới nhất (cài đè, không cần gỡ trước) → mở lại Excel. Nếu vẫn thấy giao diện cũ: đóng Excel, xóa nội dung thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, mở lại Excel. |

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
| Cài xong không thấy tab **Nhật ký thi công** | Đóng Excel → **Start menu → Nhật ký thi công (Excel)**; hoặc **Insert → My Add-ins → Developer Add-ins** (mục 2). Vẫn không có → dùng cách cài mục 5 và báo lại cho người phát triển (kèm phiên bản Excel: **File → Account → About Excel**). |
| Không thấy thẻ **SHARED FOLDER** (cách cài mục 5) | Đã tích **Show in Menu** và đã khởi động lại Excel; đường dẫn phải là đường dẫn mạng `\\...`. |
| Bấm nút trên tab không thấy gì / hộp thoại trắng | Kiểm tra Internet; thử mở <https://khoacrab-97.github.io/Addin_excel_NKTC/> bằng trình duyệt. Nếu báo "Đang có một hộp thoại khác của add-in mở" → đóng hộp thoại đó. |
| Vẫn thấy khung bên phải kiểu cũ (bản ≤ 1.2.1) | Chạy lại bộ cài 1.3.0 trở lên; nếu vẫn còn, xóa nội dung `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` rồi mở lại Excel. |
| Đã cập nhật nhưng vẫn thấy bản cũ | Đóng Excel → xóa toàn bộ nội dung thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` → mở lại Excel. |
| Nút trên giao diện không có biểu tượng | Mạng chặn CDN biểu tượng của Microsoft; chức năng vẫn dùng bình thường (các nút điều hướng ngày dùng ký tự « ‹ › »). |

## 8. Dành cho người phát triển

- Chạy bản phát triển: `npm start` (sideload manifest `localhost`). **Không** dùng đồng thời với bản đã cài bằng bộ cài — hai bản có cùng ID add-in; gỡ bản cài trước (mục 4).
- Bộ cài: `installer/NhatKyThiCong.nsi` (NSIS 3). Build: `npm run build` rồi `npm run build:installer` → `dist/NhatKyThiCong_Setup.exe` (cần `makensis`). CI và workflow phát hành tự build.
- Phát hành: workflow `.github/workflows/deploy-pages.yml` chạy khi merge vào `main` (hoặc chạy tay trong tab **Actions**): đưa add-in + bộ cài lên GitHub Pages, rồi tạo GitHub Release `v<phiên bản>` kèm `NhatKyThiCong_Setup_<phiên bản>.exe` (bỏ qua nếu phiên bản đã có Release — muốn phát hành Release mới phải tăng `version` trong `package.json`). Yêu cầu một lần: **Settings → Pages → Source: GitHub Actions**.
