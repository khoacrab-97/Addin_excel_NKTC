# MASTER_CONTEXT — ConstructionDailyLog

> Tài liệu nền tảng của dự án. Mọi task phải đọc file này trước khi triển khai (xem `CLAUDE.md` §2).
> Đặc tả nghiệp vụ chi tiết: [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md). Tiến độ: [`ROADMAP.md`](ROADMAP.md).

## 1. Mục tiêu sản phẩm

Excel Add-in (VBA, `.xlam`, tên tab **Nhật ký TC**) hỗ trợ lập **Nhật ký thi công (NKTC)** từ danh sách công việc có thời gian thực hiện, gồm:

| Mã | Nhóm chức năng | Mô tả ngắn |
|---|---|---|
| F1 | Xuất nhật ký theo ngày | Trải danh sách công việc theo từng ngày → sheet mới gồm 2 cột **Ngày** và **Nội dung**. |
| F2 | Xuất theo form mẫu (Template Mapping) | Chèn form mẫu dựng sẵn; hộp thoại "Xuất vào form": chọn vùng nguồn (sheet F1 đã thêm cột phụ, dòng đầu là tiêu đề), gán từng cột vào ô của form (+ Thêm / Delete), Gán → form hiện ngày đầu; điều hướng ngày để hiển thị nhật ký từng ngày lên form. |
| F3 | Tiện ích | Hướng dẫn sử dụng, kiểm tra phiên bản, cập nhật. |
| F4 | Tạo bảng nhập tiến độ | Tạo Excel Table **STT \| Nội dung công việc \| Bắt đầu \| Kết thúc** (sheet mới hoặc vị trí chọn trên sheet hiện tại, mặc định 10 dòng) để người dùng tự điền rồi xuất bằng F1. |

## 2. Người dùng

- Cán bộ kỹ thuật, chỉ huy trưởng công trường, QS/QA-QC lập hồ sơ nhật ký thi công.
- Môi trường chính: **Microsoft Excel Desktop trên Windows**.

## 3. Dữ liệu đầu vào (Input)

Danh sách nhật ký dạng bảng (có dòng tiêu đề), mỗi dòng là một công việc:

| Nhóm | Trường | Bắt buộc |
|---|---|---|
| Chính | Nội dung công việc | Có |
| Chính | Ngày bắt đầu | Có |
| Chính | Ngày kết thúc | Có (xem quy tắc mặc định trong BUSINESS_SPEC) |
| Phụ | Thời tiết | Không |
| Phụ | Ban chỉ huy – CBKT | Không |
| Phụ | Số lượng công nhân | Không |
| Phụ | Số lượng thiết bị | Không |
| Phụ | Nội dung nghiệm thu | Không |

## 4. Kiến trúc kỹ thuật

Từ 3.0.0 add-in là **VBA (`.xlam`)**; bản Office.js (1.x – 2.0.1) đã gỡ khỏi repo ngày 10/11/2026 (ROADMAP Phase 13, mã cũ còn trong lịch sử Git, tag `v2.0.1`).

| Thành phần | Lựa chọn | Ghi chú |
|---|---|---|
| Nền tảng | Excel add-in VBA: bộ nạp `NhatKyThiCong.xlam` + lõi `NhatKyThiCong_Core.xlam`; tab Ribbon **Nhật ký TC** (customUI14) luôn hiện | Không cần Internet khi dùng. |
| Phiên bản Excel tối thiểu | Excel 2016 / 2019 / 2021 / 2024 / Microsoft 365 trên Windows, 32 và 64-bit (`#If VBA7`) | |
| Hộp thoại | UserForm sinh khi build từ `vba/forms/*.json` + `.frmcode`; chọn vùng bằng `Application.InputBox(Type:=8)` | Không dùng ActiveX/RefEdit (Microsoft 365 chặn). Thông báo/tiêu đề tiếng Việt qua `MessageBoxW` / `DefWindowProcW`. |
| Mã nguồn | `vba/src`, `vba/loader` viết UTF-8; `vba/tools/vba_ascii.py` đổi sang ASCII (`U("\uXXXX")`) trong `vba/build/*-ascii` | CI kiểm tra mã ASCII khớp nguồn. |
| Build | `vba/build/build.cmd` trên Windows có Excel (COM): nạp mã, tạo UserForm, chép biểu mẫu, biên dịch, gắn Ribbon, kiểm thử, ghi `version.txt` | Cần bật "Trust access to the VBA project object model". |
| Test | `modNKTC_Tests.NKTC_RunTests` (chạy khi build và trong **Phiên bản**); unit test Python cho công cụ chuyển mã | Bắt buộc cho phần lõi. |
| Cấu hình gán (R2.8) | CustomXMLPart `urn:nhatkytc:config:1` trong workbook | |
| Cập nhật | Lõi tự kiểm tra `version.txt` trên GitHub Pages `/vba/` mỗi ngày; kiểm tra kích thước, ZIP, SHA-256; bộ nạp thay lõi khi mở Excel | |
| Phân phối | Bộ cài NSIS `NhatKyTC_Setup_<phiên bản>.exe` (per-user, khóa `OPEN` của Excel) qua GitHub Release `vba-v<phiên bản>` | |

### Phân tầng mã nguồn (`vba/src`)

```
modNKTC_Text, modNKTC_Date, modNKTC_DailyLog, modNKTC_Mapping, modNKTC_Config
             Phần lõi: không gọi đối tượng Excel (dữ liệu ô dạng Range.Value2), có kiểm thử
modNKTC_Excel        Lớp Excel: đọc vùng, tạo sheet/bảng
modNKTC_Commands, modNKTC_FormCommands
             Lệnh của các nút Ribbon (F1, F2, F4), lưu/đọc cấu hình
modNKTC_UI, modNKTC_Ribbon, modNKTC_Update, frm*.frmcode
             Thông báo, Ribbon, tự cập nhật, hộp thoại
vba/loader/modNKTC_Loader   Bộ nạp: thay lõi đã tải (kiểm lại SHA-256) rồi mở lõi
```

Nguyên tắc: lệnh gọi lớp Excel (I/O) và phần lõi (xử lý); phần lõi không gọi đối tượng Excel.

## 5. Ràng buộc an toàn dữ liệu Excel (tóm tắt từ `CLAUDE.md` §4)

- F1 luôn ghi ra **sheet mới**, không ghi đè sheet đã có.
- F2 chỉ ghi vào **các ô đích đã cấu hình**; chỉ ghi giá trị (values), giữ nguyên định dạng; không ghi đè ô chứa công thức; xác nhận trước khi ghi vào ô đích đang có dữ liệu.
- Mỗi lần xuất/điều hướng đều đọc lại dữ liệu nguồn mới nhất.

## 6. Thuật ngữ

| Thuật ngữ | Ý nghĩa |
|---|---|
| NKTC | Nhật ký thi công |
| Sheet theo ngày (Daily sheet) | Kết quả F1: mỗi dòng một ngày |
| Form mẫu (Template) | Sheet biểu mẫu nhật ký thi công của dự án/chủ đầu tư |
| Mapping | Cặp (cột nguồn → ô đích trên form mẫu) |
| CBKT | Cán bộ kỹ thuật |
