# Kiểm thử nghiệm thu trên Excel Desktop Windows

> Mục đích: xác nhận **Verified** cho P0.6, P1.6, P2.8, P4.2 (ROADMAP). Thực hiện trên máy thật sau khi cài theo `docs/HUONG_DAN_CAI_DAT.md`.
> Dữ liệu: `test/fixtures/test_addin_xuat_nktc.xlsx` (tiến độ), `test/fixtures/demo_f2.xlsx` (nhật ký theo ngày + biểu mẫu), `templates/Mau_NKTC.xlsx`.
> Ghi kết quả: Đạt / Không đạt + ghi chú (ảnh chụp màn hình nếu lỗi).

## Thông tin máy kiểm thử

| Mục | Giá trị |
|---|---|
| Phiên bản Office (**Tiện ích ▾ → Phiên bản**) | |
| ExcelApi tối đa (**Tiện ích ▾ → Phiên bản**) | |
| Windows | |
| Ngày kiểm thử / người kiểm thử | |

## A. Cài đặt (P0.6, P4.6)

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| A1 | Đóng Excel → chạy `NhatKyThiCong_Setup_<phiên bản>.exe` → giữ tích "Mở Excel và kích hoạt add-in" → Hoàn tất | Excel mở, khung chào mừng hiện một lần; Ribbon có tab **Nhật ký thi công** | |
| A2 | Đóng khung chào mừng; đóng Excel, mở một file bất kỳ | Tab **Nhật ký thi công** vẫn còn, có 4 nhóm: Tiến độ / Form mẫu / Chuyển ngày / Tiện ích (menu **Tiện ích ▾** gồm Hướng dẫn, Phiên bản); **không** có khung bên phải | |
| A3 | **Tiện ích ▾** → **Phiên bản** | Hộp thoại: phiên bản add-in = phiên bản trên trang phát hành; "Bạn đang dùng phiên bản mới nhất" | |
| A4 | Nút **Hướng dẫn** | Hộp thoại hướng dẫn; link tải biểu mẫu mở được | |
| A5 | Cài đè bản mới; gỡ qua Settings → Apps | Cài đè không lỗi; gỡ xong tab Ribbon biến mất sau khi mở lại Excel | |

## A'. F4 – Tạo bảng (P7.4)

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| T1 | File mới → **Tạo bảng** → giữ "Sheet mới", 10 dòng → Tạo bảng | Sheet `TienDo`, bảng A1:D11 có màu, tiêu đề STT/Nội dung công việc/Bắt đầu/Kết thúc, STT 1..10, con trỏ ở B2 | |
| T2 | Gõ 2 công việc (ngày dd/mm/yyyy); gõ thêm ở dòng 12 | Ngày hiển thị dd/mm/yyyy; dòng 12 tự vào bảng | |
| T3 | Sheet khác: chọn ô C5 → **Tạo bảng** → "Sheet hiện tại", 5 dòng | Bảng C5:F10 | |
| T4 | **Tạo bảng** lần nữa tại C5 (vùng đã có bảng) | Báo vùng đang có dữ liệu, không ghi | |
| T5 | Đặt con trỏ trong bảng T2 → **Xuất theo ngày** | 3 cột tự nhận diện; xuất đúng các ngày | |

## B. F1 – Xuất nhật ký theo ngày (P1.6) — file `test_addin_xuat_nktc.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| B1 | Sheet TDTC: đặt con trỏ 1 ô trong bảng → **Xuất theo ngày** | Hộp thoại: ô Vùng dữ liệu `TDTC!A1:C23`, 3 cột "Cột A – …", "Cột B – …", "Cột C – …" | |
| B1b | Bấm ⬆ ở ô Vùng dữ liệu → hộp thoại đóng, hiện thanh chọn "Xuất nhật ký theo ngày – Vùng dữ liệu" → kéo chọn `A1:C10` trên sheet → bấm ⬇ ở cuối thanh | Trong lúc kéo chọn, ô nhập của thanh hiện ngay `TDTC!A1:C10` và vùng đó có viền nét đứt nháy;  bấm ⬇ → hộp thoại ban đầu mở lại, ô Vùng dữ liệu = `TDTC!A1:C10`, 3 cột giữ nguyên. Lặp lại nhưng đóng thanh bằng X → giữ vùng cũ | |
| B2 | **Xuất nhật ký** | Sheet mới `NK_TheoNgay`, 200 dòng + tiêu đề, cột Ngày dd/mm/yyyy; hộp thoại kết quả "200 ngày, từ 15/06/2026 đến 31/12/2026", "22 công việc hợp lệ", xem trước 5 ngày | |
| B3 | Kiểm tra ngày 06/08/2026 | Ô Nội dung có 13 dòng, giữ nguyên nội dung (không tự thêm "- "), **không** báo lỗi công thức (#NAME?) | |
| B4 | Xuất lần 2 | Tạo `NK_TheoNgay_2`, sheet cũ giữ nguyên | |
| B5 | Chọn nguyên cột `A:C` rồi xuất; bấm **Hủy** ở hộp thoại | Kết quả giống B2; Hủy thì không tạo sheet | |

## C. F2 – Xuất theo form mẫu (P2.8) — file `demo_f2.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| C0 | File chưa có biểu mẫu → **Chèn form mẫu** (Microsoft 365 / Excel 2021+) | Thêm sheet "Mẫu NKTC" ở cuối; bấm lần 2 → báo đã có. Excel cũ → hướng dẫn tải file mẫu | |
| C1 | **Xuất vào form** | Vùng nguồn tự nhận diện `NK_TheoNgay!A1:I201`; Cột ngày = Cột A – Ngày; 200 ngày | |
| C2 | Gán theo bảng BUSINESS_SPEC §F2.3: mỗi dòng chọn cột (trái); bên phải bấm ⬆ → thanh chọn → bấm ô merge trên "Mẫu NKTC" (thanh hiện `Mẫu NKTC!E2`…, viền nháy quanh cả ô gộp E2:I2) → ⬇; + Thêm; chọn dòng thừa → Delete → **Gán** | Mỗi lần bấm ⬇ quay lại đúng hộp thoại Xuất vào form, viền nháy biến mất (không còn hình nào trên sheet), các dòng đã gán giữ nguyên; ô gộp lấy ô trên-trái; dòng 1 (Cột A – Ngày → `Mẫu NKTC!E2`) → ô E2 = "Ngày 15 tháng 06 năm 2026" | |
| C2b | Thử vùng nguồn có dòng tiêu đề bị gộp (merge) | Báo không chọn được, phải gỡ merge | |
| C3 | **Ngày sau** rồi **Ngày đầu** | Biểu mẫu hiển thị "Ngày 15 tháng 06 năm 2026", thời tiết, nhân lực, nội dung | |
| C4 | Ô nội dung 3.1 | Hiển thị đúng các dòng "- …", **không** thấy dấu `'` ở đầu (R2.11) | |
| C5 | Ô tích thời tiết | Tích đúng theo ghi chú (Nắng / Mưa / Bình thường) | |
| C6 | Nhóm **Chuyển ngày**: kiểm tra 4 nút Ngày đầu · Ngày trước · Ngày sau · Ngày cuối cùng cỡ icon, ngang hàng; bấm **Ngày cuối**, **Ngày trước**, **Ngày sau** | Dòng ngày của biểu mẫu = "Ngày dd tháng mm năm yyyy"; Biểu mẫu đổi theo ngày, không hiện hộp thoại; mục 4, 5 có nội dung sự cố / kiến nghị; ở ngày cuối bấm Ngày sau → báo "Đang ở ngày cuối cùng" | |
| C7 | Sửa ô Thời tiết trong NK_TheoNgay → **Ngày sau** rồi **Ngày trước** | Biểu mẫu nhận giá trị mới | |
| C8 | Lưu file, đóng, mở lại → **Ngày sau** | Cấu hình gán cột còn nguyên; ghi luôn, không hỏi lại | |
| C9 | **File → Print** (Xem trước bản in) | A4 đứng, 2 trang, ngắt trước mục 4 | |
| C10 | Mở bằng Excel 2016 (nếu có) | C1–C9 đạt | |
| C11 | Sau khi đã gán (C2): trong NK_TheoNgay chèn 1 cột mới giữa cột B và C (đặt tiêu đề), thêm 1 dòng ngày mới ở cuối → bấm **Ngày sau**, **Ngày cuối** | Không cần chọn lại vùng / gán lại; Thời tiết, nhân lực… vẫn hiện đúng cột; **Ngày cuối** hiện ngày vừa thêm | |
| C12 | Đổi tên tiêu đề cột "Thời tiết" thành "Thời tiết (sáng)" → **Ngày trước**; mở **Xuất vào form** | Báo không tìm thấy cột "Thời tiết" (gán vào Mẫu NKTC!C6), biểu mẫu không đổi; hộp thoại báo dòng 3 cần chọn lại cột. Đổi lại tên (khác hoa/thường vẫn được) → chạy bình thường | |

## D. Kết luận

| Hạng mục | Kết quả |
|---|---|
| P0.6 – Chạy add-in trên Excel Desktop | |
| P1.6 – F1 trên Excel thật | |
| P2.8 – F2 trên Excel thật | |
| P4.2 – Kiểm thử tích hợp với dữ liệu thực tế | |
