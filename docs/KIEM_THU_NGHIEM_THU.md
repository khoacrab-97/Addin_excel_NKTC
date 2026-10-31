# Kiểm thử nghiệm thu trên Excel Desktop Windows

> Mục đích: xác nhận **Verified** cho P0.6, P1.6, P2.8, P4.2 (ROADMAP). Thực hiện trên máy thật sau khi cài theo `docs/HUONG_DAN_CAI_DAT.md`.
> Dữ liệu: `test/fixtures/test_addin_xuat_nktc.xlsx` (tiến độ), `test/fixtures/demo_f2.xlsx` (nhật ký theo ngày + biểu mẫu), `templates/Mau_NKTC.xlsx`.
> Ghi kết quả: Đạt / Không đạt + ghi chú (ảnh chụp màn hình nếu lỗi).

## Thông tin máy kiểm thử

| Mục | Giá trị |
|---|---|
| Phiên bản Office (tab Tiện ích) | |
| ExcelApi tối đa (tab Tiện ích) | |
| Windows | |
| Ngày kiểm thử / người kiểm thử | |

## A. Cài đặt (P0.6, P4.1)

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| A1 | Cài theo hướng dẫn mục 2–4 | Nút **Nhật ký thi công** trên tab Home | |
| A2 | Mở khung add-in | 3 tab: Theo ngày / Form mẫu / Tiện ích; không báo lỗi | |
| A3 | Tab Tiện ích | Hiển thị phiên bản add-in = phiên bản trên trang phát hành | |

## B. F1 – Xuất nhật ký theo ngày (P1.6) — file `test_addin_xuat_nktc.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| B1 | Sheet TDTC: chọn `A1:C23` → **Lấy vùng đang chọn** | Nhận 3 cột (A, B, C); "22 công việc hợp lệ · từ 15/06/2026 đến 31/12/2026 · 200 ngày" | |
| B2 | **Xuất nhật ký theo ngày** | Sheet mới `NK_TheoNgay`, 200 dòng + tiêu đề; cột Ngày dạng dd/mm/yyyy | |
| B3 | Kiểm tra ngày 06/08/2026 | Ô Nội dung có 13 dòng, mỗi dòng bắt đầu "- ", **không** báo lỗi công thức (#NAME?) | |
| B4 | Xuất lần 2 | Tạo `NK_TheoNgay_2`, sheet cũ giữ nguyên | |
| B5 | Chọn nguyên cột `A:C` rồi xuất | Kết quả giống B2 | |

## C. F2 – Xuất theo form mẫu (P2.8) — file `demo_f2.xlsx`

| # | Bước | Kết quả mong đợi | Kết quả |
|---|---|---|---|
| C1 | Sheet NK_TheoNgay: chọn `A1:I201` → tab Form mẫu → **Lấy vùng đang chọn** | "9 cột · 200 ngày"; cột Ngày = A | |
| C2 | Gán cột theo bảng BUSINESS_SPEC §F2.3 (chọn ô merge trên "Mẫu NKTC" → **Ô đang chọn**) | Hiện "→ Mẫu NKTC!E2", "→ Mẫu NKTC!A21"…; ô merge lấy ô trên-trái | |
| C3 | **Xuất lên form** | Biểu mẫu hiển thị "Ngày 15 tháng 06 năm 2026", thời tiết, nhân lực, nội dung | |
| C4 | Ô nội dung 3.1 | Hiển thị đúng các dòng "- …", **không** thấy dấu `'` ở đầu (R2.11) | |
| C5 | Ô tích thời tiết | Tích đúng theo ghi chú (Nắng / Mưa / Bình thường) | |
| C6 | « ‹ › » và **Đi tới** 05/08/2026 | Biểu mẫu đổi theo ngày; mục 4, 5 có nội dung sự cố / kiến nghị | |
| C7 | Sửa ô Thời tiết trong NK_TheoNgay → bấm › rồi ‹ | Biểu mẫu nhận giá trị mới | |
| C8 | Lưu file, đóng, mở lại → tab Form mẫu | Cấu hình gán cột còn nguyên; **Xuất lại** hỏi xác nhận ghi đè | |
| C9 | **File → Print** (Xem trước bản in) | A4 đứng, 2 trang, ngắt trước mục 4 | |
| C10 | Mở bằng Excel 2016 (nếu có) | C1–C9 đạt | |

## D. Kết luận

| Hạng mục | Kết quả |
|---|---|
| P0.6 – Chạy add-in trên Excel Desktop | |
| P1.6 – F1 trên Excel thật | |
| P2.8 – F2 trên Excel thật | |
| P4.2 – Kiểm thử tích hợp với dữ liệu thực tế | |
