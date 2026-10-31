# MASTER_CONTEXT — ConstructionDailyLog

> Tài liệu nền tảng của dự án. Mọi task phải đọc file này trước khi triển khai (xem `CLAUDE.md` §2).
> Đặc tả nghiệp vụ chi tiết: [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md). Tiến độ: [`ROADMAP.md`](ROADMAP.md).

## 1. Mục tiêu sản phẩm

Excel Add-in (Office.js) hỗ trợ lập **Nhật ký thi công (NKTC)** từ danh sách công việc có thời gian thực hiện, gồm:

| Mã | Nhóm chức năng | Mô tả ngắn |
|---|---|---|
| F1 | Xuất nhật ký theo ngày | Trải danh sách công việc theo từng ngày → sheet mới gồm 2 cột **Ngày** và **Nội dung**. |
| F2 | Xuất theo form mẫu (Template Mapping) | Gán từng cột của vùng nguồn (thường là sheet F1 đã bổ sung cột nội dung phụ) vào ô của form mẫu; điều hướng tăng/giảm ngày để hiển thị nhật ký từng ngày lên form. |
| F3 | Tiện ích | Hướng dẫn sử dụng, kiểm tra phiên bản, cập nhật. |

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

## 4. Kiến trúc kỹ thuật (đề xuất)

| Thành phần | Lựa chọn | Ghi chú |
|---|---|---|
| Nền tảng | Office Add-in, Task Pane | Manifest XML (tương thích Excel Desktop Windows). |
| UI | React + TypeScript (strict) + Fluent UI React v9 | Bộ UI chuẩn của Microsoft cho Office Add-in. |
| Excel API | Office.js – `Excel.run`, ExcelApi requirement set ≥ 1.7 | Không dùng API ngoài requirement set đã khai báo. |
| Build | Webpack (theo template Yeoman `generator-office`) | |
| Test | Vitest (unit test cho tầng `core`) | |
| Lint | ESLint + Prettier | |

### Phân tầng mã nguồn

```
src/
  core/        Business logic thuần TypeScript, KHÔNG phụ thuộc Office.js (unit test được)
    dates/       Chuyển đổi Excel serial date ↔ Date, kiểm tra hợp lệ
    dailyLog/    F1: trải công việc theo ngày, gộp nội dung
    mapping/     F2: mô hình mapping cột → ô, điều hướng ngày
    validation/  Kiểm tra dữ liệu nhập
  excel/       Adapter Office.js: đọc/ghi range, tạo sheet, lưu settings
  taskpane/    React UI (các màn hình F1, F2, F3)
  version/     Thông tin phiên bản, kiểm tra cập nhật
```

Nguyên tắc: `taskpane` → gọi `excel` (I/O) và `core` (xử lý); `core` không import `excel` hay `Office`.

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
