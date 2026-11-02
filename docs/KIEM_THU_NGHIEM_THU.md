# Kiểm thử nghiệm thu trên Excel Desktop Windows

> Mục đích: xác nhận **Verified** cho P0.6, P1.6, P2.8, P4.2 (ROADMAP). Thực hiện trên máy thật sau khi cài theo `docs/HUONG_DAN_CAI_DAT.md`.
> Dữ liệu: `test/fixtures/test_addin_xuat_nktc.xlsx` (tiến độ), `test/fixtures/demo_f2.xlsx` (nhật ký theo ngày + biểu mẫu), `templates/Mau_NKTC.xlsx`.
> Ghi kết quả: Đạt / Không đạt + ghi chú (ảnh chụp màn hình nếu lỗi).

## Thông tin máy kiểm thử

| Mục | Giá trị |
|---|---|
| Phiên bản Office (nút **Phiên bản**) | |
| ExcelApi tối đa (nút **Phiên bản**) | |
| Windows | |
| Ngày kiểm thử / người kiểm thử | |

## A. Cài đặt (P0.6, P4.6)

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| A1 | Đóng Excel → chạy `NhatKyThiCong_Setup_<phiên bản>.exe` → giữ tích "Mở Excel và kích hoạt add-in" → Hoàn tất | Excel mở, khung chào mừng hiện một lần; Ribbon có tab **Nhật ký thi công** | |
| A2 | Đóng khung chào mừng; đóng Excel, mở một file bất kỳ | Tab **Nhật ký thi công** vẫn còn, có 4 nhóm: Xuất nhật ký / Form mẫu / Chuyển ngày / Tiện ích; **không** có khung bên phải | |
| A3 | Nút **Phiên bản** | Hộp thoại: phiên bản add-in = phiên bản trên trang phát hành; "Bạn đang dùng phiên bản mới nhất" | |
| A4 | Nút **Hướng dẫn** | Hộp thoại hướng dẫn; link tải biểu mẫu mở được | |
| A5 | Cài đè bản mới; gỡ qua Settings → Apps | Cài đè không lỗi; gỡ xong tab Ribbon biến mất sau khi mở lại Excel | |

## B. F1 – Xuất nhật ký theo ngày (P1.6) — file `test_addin_xuat_nktc.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| B1 | Sheet TDTC: đặt con trỏ 1 ô trong bảng → **Xuất theo ngày** | Hộp thoại: vùng `TDTC!A1:C23`, 3 cột nhận diện A, B, C | |
| B2 | **Xuất nhật ký** | Sheet mới `NK_TheoNgay`, 200 dòng + tiêu đề, cột Ngày dd/mm/yyyy; hộp thoại kết quả "200 ngày, từ 15/06/2026 đến 31/12/2026", "22 công việc hợp lệ", xem trước 5 ngày | |
| B3 | Kiểm tra ngày 06/08/2026 | Ô Nội dung có 13 dòng, mỗi dòng bắt đầu "- ", **không** báo lỗi công thức (#NAME?) | |
| B4 | Xuất lần 2 | Tạo `NK_TheoNgay_2`, sheet cũ giữ nguyên | |
| B5 | Chọn nguyên cột `A:C` rồi xuất; bấm **Hủy** ở hộp thoại | Kết quả giống B2; Hủy thì không tạo sheet | |

## C. F2 – Xuất theo form mẫu (P2.8) — file `demo_f2.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| C1 | Sheet NK_TheoNgay: đặt con trỏ trong bảng → **Vùng nguồn** | Thông báo "NK_TheoNgay!A1:I201", "Cột Ngày: A – Ngày", "200 ngày hợp lệ" | |
| C2 | Gán cột theo bảng BUSINESS_SPEC §F2.3: chọn ô merge trên "Mẫu NKTC" → **Gán cột** → chọn cột → Gán | Hộp thoại hiện đúng ô (ô merge lấy ô trên-trái); **Danh sách gán** liệt kê 9 ô | |
| C3 | **Ngày đầu** | Biểu mẫu hiển thị "Ngày 15 tháng 06 năm 2026", thời tiết, nhân lực, nội dung (nếu ô đích đã có dữ liệu → hộp thoại xác nhận ghi đè) | |
| C4 | Ô nội dung 3.1 | Hiển thị đúng các dòng "- …", **không** thấy dấu `'` ở đầu (R2.11) | |
| C5 | Ô tích thời tiết | Tích đúng theo ghi chú (Nắng / Mưa / Bình thường) | |
| C6 | **Ngày sau**, **Ngày trước**, **Ngày cuối**, **Đến ngày…** 05/08/2026 | Biểu mẫu đổi theo ngày, không hiện hộp thoại; mục 4, 5 có nội dung sự cố / kiến nghị; ở ngày cuối bấm Ngày sau → báo "Đang ở ngày cuối cùng" | |
| C7 | Sửa ô Thời tiết trong NK_TheoNgay → **Ngày sau** rồi **Ngày trước** | Biểu mẫu nhận giá trị mới | |
| C8 | Lưu file, đóng, mở lại → **Ngày sau** | Cấu hình gán cột còn nguyên; ghi luôn, không hỏi lại | |
| C9 | **File → Print** (Xem trước bản in) | A4 đứng, 2 trang, ngắt trước mục 4 | |
| C10 | Mở bằng Excel 2016 (nếu có) | C1–C9 đạt | |

## D. Kết luận

| Hạng mục | Kết quả |
|---|---|
| P0.6 – Chạy add-in trên Excel Desktop | |
| P1.6 – F1 trên Excel thật | |
| P2.8 – F2 trên Excel thật | |
| P4.2 – Kiểm thử tích hợp với dữ liệu thực tế | |
