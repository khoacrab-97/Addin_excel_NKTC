# Hướng dẫn cài đặt Add-in "Nhật ký thi công"

> Áp dụng: Excel Desktop trên Windows — Microsoft 365 hoặc Excel 2016 trở về sau.
> Địa chỉ phát hành: <https://khoacrab-97.github.io/Addin_excel_NKTC/>

## 1. Điều kiện

| Yêu cầu | Ghi chú |
|---|---|
| Excel Desktop Windows (Microsoft 365 / 2016 / 2019 / 2021 / 2024) | Excel 2016 bản mua một lần chạy được nhưng không tự đặt độ rộng cột (xem mục 6). |
| Kết nối Internet khi mở add-in | Add-in tải giao diện từ `khoacrab-97.github.io` và thư viện Office từ `appsforoffice.microsoft.com`. |
| Một thư mục chia sẻ (đường dẫn mạng dạng `\\TEN-MAY\TEN-THU-MUC`) | Excel chỉ nhận danh mục add-in từ **đường dẫn mạng**, không nhận đường dẫn ổ đĩa như `C:\...`. |

## 2. Tạo thư mục chia sẻ và tải manifest

1. Tạo thư mục, ví dụ `C:\NKTC_Addin`.
2. Chuột phải thư mục → **Properties** → thẻ **Sharing** → **Share…** → chọn tài khoản của mình (hoặc nhóm người dùng cần cài) → **Share**.
3. Ghi lại **đường dẫn mạng** hiển thị, ví dụ `\\DESKTOP-ABC123\NKTC_Addin`.
4. Mở <https://khoacrab-97.github.io/Addin_excel_NKTC/> → **Tải manifest.xml** → lưu file vào thư mục trên.

> Dùng chung cho cả nhóm: đặt thư mục trên máy chủ / NAS nội bộ, mọi người dùng cùng một đường dẫn mạng.

## 3. Khai báo thư mục trong Excel (làm một lần trên mỗi máy)

1. Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings…** → **Trusted Add-in Catalogs**.
2. Ô **Catalog Url**: dán đường dẫn mạng ở bước 2.3 → **Add catalog**.
3. Tích ô **Show in Menu** của dòng vừa thêm → **OK** → **OK**.
4. **Đóng hẳn Excel và mở lại.**

## 4. Thêm add-in vào Excel

1. Mở file Excel cần dùng.
2. **Insert** (Chèn) → **My Add-ins** / **Get Add-ins** (Excel 2016: **Insert → My Add-ins**).
3. Chọn thẻ **SHARED FOLDER** → chọn **Nhật ký thi công** → **Add**.
4. Trên tab **Home** xuất hiện nhóm **Nhật ký thi công** → bấm nút để mở khung làm việc.
5. Kiểm tra: tab **Tiện ích** hiển thị phiên bản add-in, phiên bản Office và mức **ExcelApi tối đa**.

## 5. Cập nhật

| Loại cập nhật | Cách nhận |
|---|---|
| Sửa lỗi / thêm chức năng (không đổi manifest) | **Tự động**: đóng và mở lại khung add-in (hoặc mở lại Excel). Khi có bản mới, add-in hiện thông báo "Đã có phiên bản mới" đầu khung; tab **Tiện ích** → **Kiểm tra cập nhật** để kiểm tra bất cứ lúc nào. |
| Có thay đổi manifest (thông báo kèm bản phát hành) | Tải lại `manifest.xml` (mục 2.4) ghi đè file cũ, rồi xóa cache (mục 7) và mở lại Excel. |

## 6. Khác biệt theo phiên bản Excel

| Chức năng | Microsoft 365 / 2019+ | Excel 2016 (mua một lần) |
|---|---|---|
| F1 – Xuất nhật ký theo ngày | Đầy đủ, tự đặt độ rộng cột | Đầy đủ; độ rộng cột của sheet kết quả chỉnh tay |
| F2 – Xuất theo form mẫu | Đầy đủ | Đầy đủ |

## 7. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Không thấy thẻ **SHARED FOLDER** | Kiểm tra mục 3: đã tích **Show in Menu** và đã khởi động lại Excel; đường dẫn phải là đường dẫn mạng `\\...`. |
| Khung add-in trắng / báo lỗi tải | Kiểm tra Internet; thử mở <https://khoacrab-97.github.io/Addin_excel_NKTC/taskpane.html> bằng trình duyệt. |
| Đã cập nhật nhưng vẫn thấy bản cũ | Đóng Excel → xóa toàn bộ nội dung thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` → mở lại Excel. |
| Nút trên giao diện không có biểu tượng | Mạng chặn CDN biểu tượng của Microsoft; chức năng vẫn dùng bình thường (các nút điều hướng ngày dùng ký tự « ‹ › »). |

## 8. Gỡ add-in

**Insert → My Add-ins** → chuột phải **Nhật ký thi công** → **Remove**; xóa đường dẫn ở **Trusted Add-in Catalogs** nếu không dùng nữa.

## 9. Dành cho người phát triển

- Chạy bản phát triển: `npm start` (sideload manifest `localhost`). **Không** dùng đồng thời với bản cài từ thư mục chia sẻ — hai bản có cùng ID add-in.
- Phát hành: workflow `.github/workflows/deploy-pages.yml` chạy khi push lên `main` (hoặc chạy tay trong tab **Actions**). Yêu cầu một lần: **Settings → Pages → Source: GitHub Actions**.
