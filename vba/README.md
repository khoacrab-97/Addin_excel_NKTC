# Nhật ký TC — bản add-in VBA (`.xlam`)

Bản chuyển đổi từ add-in Office.js sang add-in VBA, cùng kiểu với TienBangChu (tab luôn hiện khi mở Excel, ô chọn vùng RefEdit gốc của Excel, không cần Internet khi dùng). Tên tab: **Nhật ký TC**. Trạng thái: **In Progress** (xem `ROADMAP.md`).

## Cấu trúc

| Thư mục / tệp | Nội dung |
|---|---|
| `src/*.bas`, `src/*.cls` | Mã nguồn VBA, viết bằng UTF-8 (tiếng Việt có dấu trong chuỗi và chú thích) |
| `tools/vba_ascii.py` | Đổi mã nguồn sang ASCII trước khi nạp vào trình soạn VBA: chuỗi tiếng Việt → `U("\uXXXX")`, chú thích bỏ dấu |
| `tools/test_vba_ascii.py` | Kiểm thử công cụ trên (`python3 -m unittest test_vba_ascii`, chạy trong `vba/tools`) |

## Module

| Module | Nội dung | Tương ứng bản Office.js |
|---|---|---|
| `modNKTC_Text` | `U()`, bỏ dấu, chuẩn hóa tiêu đề | `src/core/text` |
| `modNKTC_Date` | Ngày (số nguyên hệ 1900), đọc ngày dạng chữ, mẫu "Ngày … tháng … năm …" | `src/core/dates` |
| `modNKTC_DailyLog` | F1: nhận diện cột, đọc công việc, trải theo ngày, tên sheet | `src/core/dailyLog` |
| `modNKTC_Mapping` | F2: gán cột, danh sách ngày, chuyển ngày, khớp cột theo tên, giá trị ghi ô | `src/core/mapping` |
| `modNKTC_Tests` | Kiểm thử phần lõi: `NKTC_RunTests()` → `PASS n` hoặc danh sách lỗi | `src/core/**/*.test.ts` |

## Quy ước

- Phần lõi (`modNKTC_Text`, `modNKTC_Date`, `modNKTC_DailyLog`, `modNKTC_Mapping`) không gọi đối tượng Excel; đọc dữ liệu ô dạng `Range.Value2`.
- `Const` không chứa chuỗi tiếng Việt (VBA không cho gọi hàm trong `Const`) — công cụ chuyển đổi báo lỗi nếu vi phạm.
- Không dùng `VBScript.RegExp` (Microsoft đang loại bỏ VBScript khỏi Windows).
