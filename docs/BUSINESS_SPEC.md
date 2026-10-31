# BUSINESS_SPEC — ConstructionDailyLog

> Trạng thái tài liệu: **v0.2** — các câu hỏi mở Q1–Q6 đã được chủ dự án chốt (31/10/2026), xem §5. Các quy tắc `[GIẢ ĐỊNH]` liên quan Q1–Q3 nay là quy tắc chính thức.

---

## F1. Xuất nhật ký theo ngày

### F1.1 Luồng sử dụng

1. Người dùng chọn sheet và vùng dữ liệu input (có dòng tiêu đề).
2. Add-in tự nhận diện cột theo tiêu đề (`Nội dung công việc`, `Ngày bắt đầu`, `Ngày kết thúc`); người dùng có thể chỉnh lại cột tương ứng.
3. Add-in kiểm tra dữ liệu, hiển thị số dòng hợp lệ / lỗi.
4. Nhấn **Xuất** → tạo sheet mới gồm 2 cột: **Ngày** | **Nội dung**.

### F1.2 Quy tắc nghiệp vụ

| Mã | Quy tắc |
|---|---|
| R1.1 | Một công việc xuất hiện ở **mọi ngày d** thỏa `Ngày bắt đầu ≤ d ≤ Ngày kết thúc` (tính cả hai đầu). |
| R1.2 | Phạm vi ngày xuất: từ `min(Ngày bắt đầu)` đến `max(Ngày kết thúc)`, liên tục từng ngày dương lịch. |
| R1.3 | `[GIẢ ĐỊNH]` Ngày không có công việc vẫn xuất một dòng với Nội dung để trống. |
| R1.4 | `[GIẢ ĐỊNH]` Nội dung của một ngày = các công việc đang thực hiện, giữ **thứ tự dòng trong input**, mỗi công việc một dòng trong ô, tiền tố `- `, phân tách bằng xuống dòng (Alt+Enter). |
| R1.5 | `[GIẢ ĐỊNH]` Thiếu Ngày kết thúc → coi bằng Ngày bắt đầu (công việc 1 ngày). |
| R1.6 | Ngày bắt đầu không hợp lệ, hoặc Ngày kết thúc < Ngày bắt đầu, hoặc Nội dung trống → dòng lỗi: **bỏ qua và báo cáo** (số dòng, lý do). |
| R1.7 | Chấp nhận ngày dạng Excel date (serial) và dạng text `dd/mm/yyyy`. |
| R1.8 | Cột Ngày ghi dưới dạng Excel date, định dạng `dd/mm/yyyy`; cột Nội dung bật Wrap Text. |
| R1.9 | Tên sheet kết quả mặc định `NK_TheoNgay`; nếu đã tồn tại → `NK_TheoNgay_2`, `_3`… **Không ghi đè sheet có sẵn.** |

### F1.3 Ví dụ

Input:

| Nội dung công việc | Ngày bắt đầu | Ngày kết thúc |
|---|---|---|
| Đào đất hố ga | 01/11/2026 | 03/11/2026 |
| Lắp đặt cống D600 | 02/11/2026 | 02/11/2026 |

Output:

| Ngày | Nội dung |
|---|---|
| 01/11/2026 | - Đào đất hố ga |
| 02/11/2026 | - Đào đất hố ga<br>- Lắp đặt cống D600 |
| 03/11/2026 | - Đào đất hố ga |

---

## F2. Xuất theo form mẫu (Template Mapping)

### F2.1 Luồng sử dụng

1. Người dùng (tùy nhu cầu) bổ sung các cột nội dung phụ vào sheet theo ngày (Thời tiết, Ban chỉ huy – CBKT, Số công nhân, Số thiết bị, Nội dung nghiệm thu…).
2. Mở màn hình **Mapping** (2 khu vực):
   - **Bên Input:** chọn sheet → chọn vùng dữ liệu → add-in hiển thị **số lượng cột** và danh sách cột (tên tiêu đề + chữ cột). Người dùng chỉ định **cột Ngày** (khóa điều hướng).
   - **Bên Output:** với mỗi cột được chọn, người dùng chọn **ô đích** trên form mẫu (chọn ô trên sheet rồi nhấn "Lấy ô đang chọn").
3. Nhấn **Xuất** → giá trị của ngày đầu tiên được điền vào các ô đích.
4. Thanh điều hướng: **◀ Ngày trước | [dd/mm/yyyy] | Ngày sau ▶** (thêm Đầu / Cuối / nhập ngày nhảy tới). Phạm vi điều hướng = các ngày có trong vùng input. Mỗi lần chuyển ngày, các ô đích được cập nhật theo dòng của ngày đó.

### F2.2 Quy tắc nghiệp vụ

| Mã | Quy tắc |
|---|---|
| R2.1 | Mỗi mapping = (cột nguồn → 1 ô đích). Một cột có thể map tới nhiều ô; một ô đích chỉ nhận một cột. |
| R2.2 | Ô đích nằm trong vùng merge → ghi vào ô đầu (top-left) của vùng merge. |
| R2.3 | Chỉ ghi **giá trị**; giữ nguyên định dạng, merge, viền của form mẫu. |
| R2.4 | Ô đích đang chứa **công thức** → cảnh báo, không ghi trừ khi người dùng xác nhận. |
| R2.5 | Lần ghi đầu tiên vào ô đích đang có dữ liệu → yêu cầu xác nhận. |
| R2.6 | Không ghi bất kỳ ô nào ngoài danh sách ô đích đã cấu hình. |
| R2.7 | Mỗi lần xuất/chuyển ngày: đọc lại vùng nguồn mới nhất (không dùng dữ liệu cache). |
| R2.8 | Cấu hình mapping được lưu trong workbook để dùng lại ở lần mở sau. |
| R2.9 | Vùng nguồn/ô đích bị xóa, đổi tên sheet hoặc không hợp lệ → báo lỗi rõ ràng, không ghi. |

---

## F3. Tiện ích

| Mã | Chức năng | Mô tả |
|---|---|---|
| F3.1 | Hướng dẫn | Màn hình hướng dẫn từng bước cho F1, F2, kèm file mẫu input. |
| F3.2 | Kiểm tra phiên bản | Hiển thị phiên bản đang chạy; so sánh với phiên bản mới nhất công bố trên máy chủ (file `version.json`). |
| F3.3 | Cập nhật | Web add-in tự cập nhật khi mã nguồn trên máy chủ hosting được cập nhật. Khi có thay đổi manifest, hiển thị hướng dẫn cài lại. |

---

## 5. Câu hỏi mở — đã chốt

| # | Câu hỏi | Quyết định |
|---|---|---|
| Q1 | Các cột nội dung phụ có nằm trong danh sách input (theo từng công việc) không? Nếu có, F1 có cần tự gộp chúng theo ngày không, hay người dùng tự nhập vào sheet theo ngày? | F1 chỉ xuất 2 cột; người dùng tự bổ sung cột phụ. |
| Q2 | Định dạng gộp nội dung trong một ngày (tiền tố `- `, đánh số 1., 2., …)? | `- ` mỗi dòng. |
| Q3 | Có xuất ngày không có công việc, ngày Chủ nhật/ngày lễ không? | Xuất tất cả ngày, kể cả ngày trống. |
| Q4 | F2 có cần **xuất hàng loạt** (mỗi ngày một sheet, hoặc in/PDF tất cả ngày) không? | Ngoài phạm vi hiện tại; xem xét ở giai đoạn sau. |
| Q5 | Cách phân phối add-in: Shared Folder Catalog (mạng nội bộ), Microsoft 365 Centralized Deployment, hay chỉ dùng cá nhân? Nơi hosting (GitHub Pages…)? | GitHub Pages + Shared Folder Catalog. |
| Q6 | Phiên bản Excel tối thiểu? | **Microsoft 365 hoặc Excel 2016 trở về sau.** Hệ quả kỹ thuật: xem `MASTER_CONTEXT.md` §4. |
