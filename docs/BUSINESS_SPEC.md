# BUSINESS_SPEC — ConstructionDailyLog

> Trạng thái tài liệu: **v0.2** — các câu hỏi mở Q1–Q6 đã được chủ dự án chốt (31/10/2026), xem §5. Các quy tắc `[GIẢ ĐỊNH]` liên quan Q1–Q3 nay là quy tắc chính thức.

---

## F1. Xuất nhật ký theo ngày

### F1.1 Luồng sử dụng

1. Người dùng bôi chọn vùng danh sách công việc trên sheet (dòng đầu là tiêu đề; được phép chọn nguyên cột), nhấn **Lấy vùng đang chọn**.
2. Add-in tự nhận diện cột theo tiêu đề (R1.12); người dùng có thể chỉnh lại cột tương ứng trong 3 ô chọn.
3. Add-in kiểm tra dữ liệu, hiển thị số công việc hợp lệ, khoảng ngày, danh sách dòng lỗi và xem trước 5 ngày đầu.
4. Nhấn **Xuất nhật ký theo ngày** → add-in đọc lại dữ liệu mới nhất của vùng; nếu dòng tiêu đề đã thay đổi thì dừng và yêu cầu kiểm tra lại cột; nếu không, tạo sheet mới gồm 2 cột: **Ngày** | **Nội dung**.

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
| R1.9 | Tên sheet kết quả mặc định `NK_TheoNgay`; nếu đã tồn tại → `NK_TheoNgay_2`, `_3`… (không phân biệt hoa/thường). **Không ghi đè sheet có sẵn.** |
| R1.10 | Khoảng thời gian xuất tối đa **3.660 ngày** (~10 năm); vượt quá → báo lỗi, không xuất (chặn lỗi nhập năm, ví dụ 2062 thay vì 2026). Nội dung một ngày vượt 32.767 ký tự (giới hạn ô Excel) → báo lỗi. |
| R1.11 | Dòng trống ở cả 3 cột chính → bỏ qua, không tính là lỗi. |
| R1.12 | Nhận diện cột (không phân biệt dấu, hoa/thường): **Nội dung** = `Nội dung công việc`, `Tên công việc`, `Công việc`, `Tên hạng mục/ công tác`, `Hạng mục`, `Công tác`, `Nội dung` hoặc tiêu đề chứa "công việc", "công tác", "hạng mục"; **Bắt đầu** = `Ngày bắt đầu`, `Bắt đầu`, `Từ ngày` hoặc chứa "bắt đầu"; **Kết thúc** = `Ngày kết thúc`, `Kết thúc`, `Đến ngày` hoặc chứa "kết thúc". |
| R1.13 | Ngày được hiểu theo hệ ngày **1900** của Excel (mặc định trên Windows). Cột Nội dung được đặt định dạng Text trước khi ghi, để nội dung bắt đầu bằng `-`, `=`, `+` không bị Excel hiểu là công thức. |

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
2. Tab **Form mẫu** gồm 3 khu vực:
   - **1. Nguồn dữ liệu (Input):** bôi chọn vùng dữ liệu (gồm dòng tiêu đề) → **Lấy vùng đang chọn** → add-in hiển thị vùng, **số lượng cột**, số ngày. **Cột Ngày** được tự nhận diện (R2.13), có thể chọn lại.
   - **2. Gán cột vào ô (Output):** mỗi cột nguồn một dòng (chữ cột + tiêu đề). Người dùng chọn ô (hoặc ô merge) trên form → nhấn **Ô đang chọn** ở cột tương ứng. Có thể gán một cột vào nhiều ô; **× Bỏ gán** để xóa.
   - **3. Xuất và chuyển ngày:** **Xuất lên form** → điền dữ liệu ngày đầu tiên (hoặc ngày đang mở lần trước) vào các ô đích.
3. Thanh điều hướng: **« Đầu | ‹ Trước | [dd/mm/yyyy · Ngày i/n] | Sau › | Cuối »** và ô **Nhảy tới ngày**. Phạm vi điều hướng = các ngày hợp lệ trong vùng nguồn, tăng dần. Mỗi lần chuyển ngày, add-in đọc lại vùng nguồn và cập nhật các ô đích.

### F2.2 Quy tắc nghiệp vụ

| Mã | Quy tắc |
|---|---|
| R2.1 | Mỗi mapping = (cột nguồn → 1 ô đích). Một cột có thể map tới nhiều ô; một ô đích chỉ nhận một cột. |
| R2.2 | Ô đích nằm trong vùng merge → ghi vào ô đầu (top-left) của vùng merge. Chọn nhiều ô → lấy ô trên-trái. Chọn nhiều vùng rời → báo lỗi. |
| R2.3 | Chỉ ghi **giá trị**; giữ nguyên định dạng, merge, viền của form mẫu. |
| R2.4 | Ô đích đang chứa **công thức** → cảnh báo, không ghi trừ khi người dùng xác nhận. |
| R2.5 | Lần ghi đầu tiên (mỗi lần mở task pane, hoặc khi đổi bộ ô đích) vào ô đích đang có dữ liệu → yêu cầu xác nhận, liệt kê các ô; **Hủy** thì không ghi ô nào. Sau khi xác nhận, chuyển ngày không hỏi lại. |
| R2.6 | Không ghi bất kỳ ô nào ngoài danh sách ô đích đã cấu hình. |
| R2.7 | Mỗi lần xuất/chuyển ngày: đọc lại vùng nguồn mới nhất (không dùng dữ liệu cache). |
| R2.8 | Cấu hình mapping được lưu trong workbook để dùng lại ở lần mở sau. |
| R2.9 | Vùng nguồn/ô đích bị xóa, đổi tên sheet hoặc không hợp lệ → báo lỗi rõ ràng, không ghi ô nào. Dòng tiêu đề vùng nguồn thay đổi so với cấu hình → dừng, yêu cầu kiểm tra lại gán cột. |
| R2.10 | Danh sách ngày điều hướng: dòng có ô Ngày trống → bỏ qua; ngày không hợp lệ → báo lỗi theo dòng; ngày trùng → báo lỗi, giữ dòng xuất hiện trước. |
| R2.11 | Giá trị văn bản được ghi kèm dấu `'` đầu chuỗi (ký tự đánh dấu văn bản của Excel, không hiển thị) để Excel không hiểu nhầm `- …`, `= …`, `01/11` thành công thức/ngày mà không phải đổi định dạng ô của form. Số (kể cả ngày dạng serial) giữ nguyên → hiển thị theo định dạng sẵn có của ô đích. |
| R2.12 | Ô đích chỉ nhận **một** ô; một ô đích không được gán cho 2 cột (cảnh báo khi gán). |
| R2.13 | Nhận diện cột Ngày: tiêu đề `Ngày` → tiêu đề chứa từ "ngày" → cột đầu tiên có toàn bộ ô không trống là ngày hợp lệ. |

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
| Q7 | Bảng tiến độ có dòng hạng mục tổng và dòng công tác chi tiết (phân cấp bằng khoảng trắng đầu dòng). F1 lấy dòng nào? (chốt 31/10/2026, sau khi chạy thử file tiến độ mẫu) | **Giữ tất cả các dòng**, kể cả dòng hạng mục tổng. Khoảng trắng đầu dòng được bỏ khi gộp nội dung. Dữ liệu thực tế sẽ không có thụt lề phân cấp. |
| Q8 | Biểu mẫu nhật ký thi công dùng cho F2? | Chưa có biểu mẫu thật; dùng biểu mẫu MẪU trong `test/fixtures/demo_f2.xlsx` để phát triển và chạy thử. Chạy lại khi có biểu mẫu thật. |
| Q6 | Phiên bản Excel tối thiểu? | **Microsoft 365 hoặc Excel 2016 trở về sau.** Hệ quả kỹ thuật: xem `MASTER_CONTEXT.md` §4. |
