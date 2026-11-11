# Kiểm thử nghiệm thu trên Excel Desktop Windows (bản VBA "Nhật ký TC")

> Mục đích: xác nhận **Verified** cho các hạng mục Phase 12 (ROADMAP) trên máy thật, sau khi cài theo `docs/HUONG_DAN_CAI_DAT.md`.
> Dữ liệu: `test/fixtures/test_addin_xuat_nktc.xlsx` (tiến độ), `test/fixtures/demo_f2.xlsx` (nhật ký theo ngày + biểu mẫu), `templates/Mau_NKTC.xlsx`.
> Ghi kết quả: Đạt / Không đạt + ghi chú (ảnh chụp màn hình nếu lỗi).

## Thông tin máy kiểm thử

| Mục | Giá trị |
|---|---|
| Phiên bản add-in / Excel (**Tiện ích ▾ → Phiên bản**) | |
| Windows, Excel 32/64-bit | |
| Ngày kiểm thử / người kiểm thử | |

## A. Cài đặt, cập nhật, gỡ

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| A1 | Đóng Excel → chạy `NhatKyTC_Setup_<phiên bản>.exe` → Next đến hết | Cài xong không cần quyền quản trị | |
| A2 | Mở Excel; đóng rồi mở lại với một file bất kỳ | Tab **Nhật ký TC** luôn hiện, 5 nhóm: Tiến độ / Form mẫu / Chuyển ngày / In / Tiện ích | |
| A3 | **Tiện ích ▾ → Phiên bản** | Hiện phiên bản, Excel, thư mục `%LOCALAPPDATA%\NhatKyTC`, kiểm thử phần lõi `PASS …` | |
| A4 | **Tiện ích ▾ → Kiểm tra cập nhật** | "Bạn đang dùng phiên bản mới nhất (…)" | |
| A5 | Khi có bản mới: mở Excel, đợi ~30 giây (hoặc bấm Kiểm tra cập nhật) → đóng, mở lại Excel | Thông báo đã tải bản mới; mở lại → **Phiên bản** hiện số mới | |
| A6 | Gỡ qua Settings → Apps → **Nhật ký TC (Excel add-in)** | Mở lại Excel không còn tab; add-in khác (nếu có) vẫn còn | |

## T. F4 – Tạo bảng

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| T1 | File mới → **Tạo bảng** → "Sheet mới", 10 dòng → Tạo bảng | Sheet `TienDo`, bảng A1:D11 có màu, tiêu đề STT / Nội dung công việc / Bắt đầu / Kết thúc, STT 1..10, con trỏ ở B2; thông báo tiếng Việt đúng dấu | |
| T2 | Gõ 2 công việc (ngày dd/mm/yyyy); gõ thêm ở dòng 12 | Ngày hiển thị dd/mm/yyyy; dòng 12 tự vào bảng | |
| T3 | Sheet khác: **Tạo bảng** → "Sheet hiện tại" → nút **Chọn** → chọn ô C5, 5 dòng | Bảng C5:F10 | |
| T4 | **Tạo bảng** lần nữa tại C5 | Báo vùng đang có dữ liệu, không ghi | |

## B. F1 – Xuất nhật ký theo ngày — file `test_addin_xuat_nktc.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| B1 | Sheet TDTC: đặt con trỏ 1 ô trong bảng → **Xuất theo ngày** | Ô Vùng dữ liệu = cả bảng; 3 cột tự nhận diện ("Cột A – …") | |
| B1b | Bấm **Chọn** → hộp chọn vùng của Excel → chọn `A1:C10` → OK | Hộp thoại hiện lại, Vùng dữ liệu = `'TDTC'!$A$1:$C$10`; Cancel → giữ vùng cũ | |
| B2 | **Xuất nhật ký** (vùng cả bảng) | Sheet mới `NK_TheoNgay`, cột Ngày dd/mm/yyyy; thông báo số ngày, khoảng ngày, số công việc hợp lệ | |
| B3 | Kiểm tra một ngày nhiều công việc | Ô Nội dung nhiều dòng, giữ nguyên nội dung, không lỗi công thức (#NAME?) | |
| B4 | Xuất lần 2 | Tạo `NK_TheoNgay_2`, sheet cũ giữ nguyên | |
| B5 | **Hủy** ở hộp thoại | Không tạo sheet | |

## C. F2 – Xuất theo form mẫu — file `demo_f2.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| C0 | File chưa có biểu mẫu → **Chèn form mẫu** | Thêm sheet "Mẫu NKTC" ở cuối (không cần Internet); bấm lần 2 → báo đã có | |
| C1 | **Xuất vào form** | Vùng nguồn tự nhận diện sheet `NK_TheoNgay`; Cột Ngày = Cột A – Ngày; "Có N ngày (từ – đến)" | |
| C2 | Mỗi cột: chọn **Cột nguồn**, **Ô trên form** (gõ hoặc **Chọn** → bấm ô gộp trên "Mẫu NKTC") → **Thêm**; chọn dòng thừa → **Xóa dòng** → **Gán** | Ô gộp lấy ô trên-trái; thông báo "Đã gán N ô…"; cột Ngày vào dòng "Ngày … tháng … năm …" → "Ngày 15 tháng 06 năm 2026" | |
| C2b | Vùng nguồn có dòng tiêu đề bị gộp (merge) | Báo không chọn được, phải gỡ merge | |
| C3 | Ô nhận dữ liệu đang có nội dung/công thức (lần gán đầu) | Hỏi trước khi ghi đè (OK / Cancel); đã đồng ý thì lần sau không hỏi lại | |
| C4 | Ô nội dung | Hiển thị đúng các dòng "- …", không thấy dấu `'` ở đầu | |
| C5 | **Ngày sau**, **Ngày trước**, **Ngày cuối**, **Ngày đầu** | Biểu mẫu đổi theo ngày; ở ngày cuối bấm Ngày sau → báo "Đang ở ngày cuối cùng" | |
| C6 | Sửa ô Thời tiết trong NK_TheoNgay → **Ngày sau** rồi **Ngày trước** | Biểu mẫu nhận giá trị mới | |
| C7 | Lưu file, đóng, mở lại → **Ngày sau** | Cấu hình gán còn nguyên | |
| C8 | Chèn 1 cột mới giữa cột B và C (có tiêu đề), thêm 1 dòng ngày ở cuối → **Ngày sau**, **Ngày cuối** | Không cần gán lại; các cột vẫn đúng; **Ngày cuối** hiện ngày vừa thêm | |
| C9 | Đổi tên tiêu đề "Thời tiết" → **Ngày trước** | Báo không tìm thấy cột "Thời tiết" (kèm ô đích), biểu mẫu không đổi; đổi lại tên (khác hoa/thường vẫn được) → chạy bình thường | |
| C10 | **File → Print** (xem trước) | A4 đứng, 2 trang | |

## P. F5 – In nhật ký

Chuẩn bị: file đã gán F2 (ví dụ `demo_f2.xlsx`), có máy in hoặc "Microsoft Print to PDF".

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| P1 | **In nhật ký** → Ngày in "đang hiển thị", Đầu ra "Máy in" (1 bản) | Máy in nhận đúng biểu mẫu của ngày đang hiển thị; Vùng in mặc định = Print Area (hoặc UsedRange) của sheet form | |
| P2 | Ngày in "một ngày" (chọn 1 ngày) → Đầu ra PDF | Hỏi nơi lưu; file `NKTC_dd-mm-yyyy.pdf` đúng ngày, nội dung đúng | |
| P3 | Ngày in "khoảng ngày" 5 ngày → PDF | File `NKTC_<từ>_den_<đến>.pdf` gồm đúng 5 ngày theo thứ tự ngày; mỗi ngày đủ số trang theo thiết lập trang (ví dụ 2 trang/ngày → 10 trang) | |
| P4 | Ngày in "tất cả" khi có hơn 50 ngày | Hỏi xác nhận; chọn đồng ý rồi bấm **Esc** giữa chừng → dừng, form trở về ngày cũ | |
| P5 | Dựng biểu mẫu riêng trên sheet khác (có công thức tham chiếu form) → **Chọn** làm Vùng in | In đúng vùng đã chọn, công thức ra giá trị đúng của từng ngày | |
| P6 | Sau khi in từ ngày cuối trở về: kiểm tra danh sách sheet và ngày đang hiển thị | Form về đúng ngày cũ; file không có thêm sheet nào (kể cả sheet tạm) | |
| P7 | Mẫu in tự dựng trên sheet riêng (ô đích của form không bị khóa), sheet mẫu in có **Protect Sheet**: (a) không mật khẩu; (b) có mật khẩu → In nhật ký | (a) In được, sheet gốc vẫn khóa; (b) báo bằng tiếng Việt "Sheet chứa vùng in đang khóa (Protect Sheet) có mật khẩu…", không in, form về ngày cũ | |
| P8 | So sánh file PDF xuất ra với **Print Preview** của file gốc | Giống nhau: số trang, phông chữ, độ rộng cột | |
| P9 | Sửa tay một ô trên form → In nhật ký, Ngày in "đang hiển thị" | Bản in có nội dung đã sửa tay; form giữ nguyên nội dung sửa tay sau khi in | |

## D. Kết luận

| Hạng mục | Kết quả |
|---|---|
| Cài đặt / tab luôn hiện / tự cập nhật / gỡ (A) | |
| F4 Tạo bảng (T) | |
| F1 Xuất theo ngày (B) | |
| F2 Xuất vào form + Chuyển ngày (C) | |
| F5 In nhật ký (P) | |
