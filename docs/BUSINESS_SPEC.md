# BUSINESS_SPEC — ConstructionDailyLog

> Trạng thái tài liệu: **v0.2** — các câu hỏi mở Q1–Q6 đã được chủ dự án chốt (31/10/2026), xem §5. Các quy tắc `[GIẢ ĐỊNH]` liên quan Q1–Q3 nay là quy tắc chính thức.

---

## F0. Giao diện (từ 1.3.0)

| Mã | Quy tắc |
|---|---|
| R0.1 | Chức năng nằm trên tab Ribbon riêng **Nhật ký thi công**: nhóm **Xuất nhật ký** (Xuất theo ngày), **Form mẫu** (Vùng nguồn, Gán cột, Danh sách gán), **Chuyển ngày** (Ngày đầu, Ngày trước, Ngày sau, Ngày cuối, Đến ngày…), **Tiện ích** (Hướng dẫn, Phiên bản). Không có khung làm việc bên phải (theo yêu cầu chủ dự án 02/11/2026). |
| R0.2 | Bước cần nhập liệu/xác nhận dùng hộp thoại (Office Dialog API 1.1 — có từ Excel 2016) mở giữa màn hình và tự đóng khi xong. Dữ liệu vào hộp thoại qua URL/localStorage (cùng domain), kết quả trả về qua `messageParent` — chỉ nhận tin có origin của add-in (khi Office cung cấp origin); hai phía kiểm tra dữ liệu nhận được, dữ liệu sai coi như Hủy. |
| R0.3 | Khung bên phải chỉ hiện **một lần** khi bộ cài kích hoạt add-in (trang chào mừng), người dùng đóng được. |

## F1. Xuất nhật ký theo ngày

### F1.1 Luồng sử dụng

> Từ 1.3.0 mọi chức năng nằm trên tab Ribbon **Nhật ký thi công** (R0.1); không còn khung làm việc bên phải.

1. Người dùng bôi chọn vùng danh sách công việc (dòng đầu là tiêu đề; được phép chọn nguyên cột) — hoặc chỉ đặt con trỏ 1 ô trong sheet tiến độ (add-in lấy toàn bộ vùng có dữ liệu của sheet) — rồi bấm nút **Xuất theo ngày**.
2. Hộp thoại hiện vùng dữ liệu và 3 ô chọn cột đã tự nhận diện theo tiêu đề (R1.12); người dùng chỉnh lại nếu cần → **Xuất nhật ký**.
3. Add-in đọc lại dữ liệu mới nhất của vùng; nếu dòng tiêu đề đã thay đổi thì dừng và yêu cầu bấm lại; nếu không, tạo sheet mới gồm 2 cột: **Ngày** | **Nội dung**.
4. Hộp thoại kết quả: tên sheet, số ngày, khoảng ngày, số công việc hợp lệ, danh sách dòng lỗi (tối đa 10) và xem trước 5 ngày đầu.

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
2. Nhóm **Form mẫu** trên Ribbon:
   - **Vùng nguồn:** bôi chọn vùng dữ liệu theo ngày (gồm tiêu đề) hoặc đặt con trỏ trong sheet đó → bấm → add-in lưu vùng nguồn, tự nhận diện **cột Ngày** (R2.13) và báo vùng, cột Ngày, số ngày hợp lệ. Nếu tiêu đề giống cấu hình cũ thì giữ các ô đã gán; khác thì bỏ gán cũ.
   - **Gán cột:** chọn ô (hoặc ô merge) trên form → bấm → hộp thoại chọn cột ghi vào ô đó. Ô đã gán: hộp thoại hiện cột hiện tại, chọn cột khác để thay hoặc **Bỏ gán ô này**. Một cột có thể gán vào nhiều ô.
   - **Danh sách gán:** bảng ô đích → cột nguồn; tích **Bỏ gán** để xóa; đổi **cột Ngày**.
3. Nhóm **Chuyển ngày**: **Ngày đầu**, **Ngày trước**, **Ngày sau**, **Ngày cuối**, **Đến ngày…** (hộp thoại nhập dd/mm/yyyy). Phạm vi điều hướng = các ngày hợp lệ trong vùng nguồn, tăng dần; ngày đang hiển thị được lưu trong cấu hình. Mỗi lần bấm, add-in đọc lại vùng nguồn và ghi các ô đích; không mở hộp thoại trừ khi cần xác nhận (R2.5), đã ở ngày đầu/cuối, hoặc có lỗi. Ngày đang xem hiển thị trên chính biểu mẫu (ô gán cột Ngày) — Ribbon Office.js không hiển thị được nhãn động.

### F2.2 Quy tắc nghiệp vụ

| Mã | Quy tắc |
|---|---|
| R2.1 | Mỗi mapping = (cột nguồn → 1 ô đích). Một cột có thể map tới nhiều ô; một ô đích chỉ nhận một cột. |
| R2.2 | Ô đích nằm trong vùng merge → ghi vào ô đầu (top-left) của vùng merge. Chọn nhiều ô → lấy ô trên-trái. Chọn nhiều vùng rời → báo lỗi. |
| R2.3 | Chỉ ghi **giá trị**; giữ nguyên định dạng, merge, viền của form mẫu. |
| R2.4 | Ô đích đang chứa **công thức** → cảnh báo, không ghi trừ khi người dùng xác nhận. |
| R2.5 | Lần ghi đầu tiên cho một bộ ô đích vào ô đang có dữ liệu/công thức → hộp thoại xác nhận liệt kê các ô; **Hủy** (hoặc đóng hộp thoại) thì không ghi ô nào. Xác nhận được lưu trong workbook gắn với đúng bộ ô đích (`confirmedKey`); chuyển ngày sau đó, kể cả khi mở lại file, không hỏi lại. Thay đổi gán cột → xác nhận cũ hết hiệu lực, hỏi lại. Xác nhận đọc từ file không khớp bộ ô đích → bị bỏ. Mỗi lần chỉ chạy một lệnh (bấm nút khác khi đang mở hộp thoại → bỏ qua). |
| R2.6 | Không ghi bất kỳ ô nào ngoài danh sách ô đích đã cấu hình. |
| R2.7 | Mỗi lần xuất/chuyển ngày: đọc lại vùng nguồn mới nhất (không dùng dữ liệu cache). |
| R2.8 | Cấu hình mapping được lưu trong workbook để dùng lại ở lần mở sau. |
| R2.9 | Vùng nguồn/ô đích bị xóa, đổi tên sheet hoặc không hợp lệ → báo lỗi rõ ràng, không ghi ô nào. Dòng tiêu đề vùng nguồn thay đổi so với cấu hình → dừng, yêu cầu kiểm tra lại gán cột. |
| R2.10 | Danh sách ngày điều hướng: dòng có ô Ngày trống → bỏ qua; ngày không hợp lệ → báo lỗi theo dòng; ngày trùng → báo lỗi, giữ dòng xuất hiện trước. |
| R2.11 | Giá trị văn bản được ghi kèm dấu `'` đầu chuỗi (ký tự đánh dấu văn bản của Excel, không hiển thị) để Excel không hiểu nhầm `- …`, `= …`, `01/11` thành công thức/ngày mà không phải đổi định dạng ô của form. Số (kể cả ngày dạng serial) giữ nguyên → hiển thị theo định dạng sẵn có của ô đích. |
| R2.12 | Ô đích chỉ nhận **một** cột: gán cột cho ô đã gán → thay cột cũ (hộp thoại hiện cột hiện tại trước khi thay). |
| R2.14 | Cấu hình đọc từ workbook vượt giới hạn → bỏ toàn bộ cấu hình (fail closed): tên sheet ≤ 31 ký tự; địa chỉ vùng nguồn phải là một vùng liên tục không kèm tên sheet (`A1:I201`, `A:C`, `3:5`); ≤ 16.384 tiêu đề, mỗi tiêu đề ≤ 32.767 ký tự; ≤ 500 ô đích; chỉ số cột < 16.384. |
| R2.13 | Nhận diện cột Ngày: tiêu đề `Ngày` → tiêu đề chứa từ "ngày" → cột đầu tiên có toàn bộ ô không trống là ngày hợp lệ. |

### F2.3 Biểu mẫu Excel "Tình hình thi công hàng ngày" (`templates/Mau_NKTC.xlsx`)

Chuyển từ mẫu Word, sheet **Mẫu NKTC**, khổ A4 đứng, lề trái 2,5 cm / phải, trên, dưới 2 cm, Times New Roman 12, in 2 trang (ngắt trang trước mục 4).

| Mục (theo mẫu Word) | Ô đích (gán bằng F2) | Ghi chú |
|---|---|---|
| Ngày… tháng… năm… | `E2` (merge E2:I2) | Nhận ngày dạng số; định dạng hiển thị `Ngày dd tháng mm năm yyyy`. |
| 1. Thời tiết – ô tích Bình thường / Mưa / Nắng | — (công thức `B5`, `E5`, `H5`) | Tự tích theo "Ghi chú thời tiết": chứa "mưa" → Mưa; chứa "nắng" → Nắng; có nội dung khác → Bình thường. |
| 1. Ghi chú thời tiết | `C6` (merge C6:I6) | |
| 1. Nhiệt độ Sáng / Trưa / Chiều / Tối | `C7`, `E7`, `G7`, `I7` | Định dạng `0"°C"`. |
| 2.1. Nhân lực – Ban chỉ huy, CBKT | `C11` (merge C11:I11) | |
| 2.1. Số lượng công nhân | `C12` (merge C12:D12) | Kèm đơn vị "người". |
| 2.1. Nhân lực – chi tiết | `A13` (merge A13:I14) | |
| 2.2. Số lượng thiết bị | `C16` (merge C16:D16) | Kèm đơn vị "chiếc". |
| 2.2. Thiết bị – chi tiết | `A17` (merge A17:I18) | |
| 3.1. Nội dung, khối lượng công việc và tiến độ | `A21` (merge A21:I38, 18 dòng) | Nhận cột Nội dung của F1. |
| 3.2. Nghiệm thu công việc xây dựng hàng ngày | `A40` (merge A40:I47) | |
| 4. Sự cố, hư hỏng, vấn đề phát sinh | `A50` (merge A50:I53) | |
| 5. Kiến nghị, ý kiến chỉ đạo | `A56` (merge A56:I59) | |
| 6. Nhận xét của kỹ sư giám sát; 7. Chỉ huy trưởng tiếp thu | — | Dòng gạch chấm để ghi tay, ký tên. |

Sinh lại biểu mẫu: `python3 -I scripts/fixtures/make_nktc_workbooks.py test/fixtures/test_addin_xuat_nktc.xlsx`.

---

## F3. Tiện ích

| Mã | Chức năng | Mô tả |
|---|---|---|
| F3.1 | Hướng dẫn | Nút **Hướng dẫn** (nhóm Tiện ích) → hộp thoại: các bước F1, F2 theo nút Ribbon; link tải `templates/Mau_NKTC.xlsx` (biểu mẫu), `templates/Mau_TienDo.xlsx` (danh sách công việc mẫu), bảng ô đích, hướng dẫn cài đặt. |
| F3.2 | Kiểm tra phiên bản | Nút **Phiên bản** → hộp thoại: phiên bản add-in, nền tảng, phiên bản Office, ExcelApi tối đa; tự kiểm tra `version.json` trên máy chủ phát hành khi mở, nút **Kiểm tra lại** (R3.1–R3.3). |
| F3.3 | Cập nhật | Mã add-in tải từ máy chủ phát hành mỗi lần Excel nạp add-in → tự cập nhật. Có bản mới hơn bản đang chạy → hộp thoại Phiên bản hiển thị các bước cập nhật. |

| Mã | Quy tắc |
|---|---|
| R3.1 | So sánh phiên bản `x.y.z` theo số (1.10.0 > 1.9.9). Bản đang chạy nhỏ hơn → "Đã có phiên bản mới"; bằng → "mới nhất"; lớn hơn → "bản phát triển". |
| R3.2 | `version.json` được kiểm tra trước khi dùng; lỗi mạng, hết thời gian chờ (10 giây), HTTP lỗi hoặc nội dung sai → "Không kiểm tra được phiên bản mới", không chặn các chức năng khác. |
| R3.3 | Yêu cầu tải `version.json` kèm tham số chống cache; dùng `XMLHttpRequest` (tương thích Internet Explorer 11). |

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
| Q8 | Biểu mẫu nhật ký thi công dùng cho F2? | Mẫu Word "ND_Nhật ký thi công" (Tình hình thi công hàng ngày, mục 1–7) do chủ dự án cung cấp ngày 31/10/2026, đã chuyển sang Excel: `templates/Mau_NKTC.xlsx` (xem §F2.3). |
| Q6 | Phiên bản Excel tối thiểu? | **Microsoft 365 hoặc Excel 2016 trở về sau.** Hệ quả kỹ thuật: xem `MASTER_CONTEXT.md` §4. |
