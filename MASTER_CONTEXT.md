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

## 4. Kiến trúc kỹ thuật

| Thành phần | Lựa chọn | Ghi chú |
|---|---|---|
| Nền tảng | Office Add-in (TaskPaneApp manifest XML) + **add-in commands**: tab Ribbon riêng **Nhật ký thi công**, nút chạy hàm trong function file (`ExecuteFunction`), hộp thoại Office Dialog API 1.1 | Từ 1.3.0 không còn khung làm việc bên phải (BUSINESS_SPEC R0.1–R0.3). |
| Phiên bản Excel tối thiểu | **Microsoft 365 hoặc Excel 2016 trở về sau** (Q6) | |
| Excel API | Office.js, requirement set **ExcelApi 1.1** làm mức tối thiểu | Excel 2016 bản mua một lần chỉ hỗ trợ ExcelApi 1.1. API cao hơn chỉ được dùng kèm kiểm tra `isSetSupported` và có phương án thay thế. |
| Trình duyệt nhúng | Phải chạy được trên **Internet Explorer 11** | Excel 2016/2019 bản mua một lần dùng IE11 để hiển thị add-in. |
| UI | **React 17** + **Fluent UI React v8** + TypeScript (strict) | React 18 và Fluent UI v9 không hỗ trợ IE11. |
| Build | Webpack 5 + Babel (`preset-env`, core-js polyfill), output **ES5** | `tsc --noEmit` dùng để kiểm tra kiểu. Môi trường dev: Node.js ≥ 22.12. |
| Hosting | GitHub Pages: `https://khoacrab-97.github.io/Addin_excel_NKTC/` (Q5) | Build production tự thay URL `localhost` trong manifest. |
| Phân phối | Bộ cài Windows (NSIS, per-user, đăng ký `HKCU\…\Wef\Developer` + workbook kích hoạt) qua GitHub Releases / Pages; Shared Folder Catalog dự phòng (Q5) | |
| Test | Vitest (`src/**/*.test.ts`, môi trường Node) | Bắt buộc cho business logic tầng `core`. |
| Lint | ESLint (typescript-eslint `strictTypeChecked`, react-hooks) + Prettier | ESLint chặn tầng `core` dùng `Office`/`Excel` hoặc import `excel`/`taskpane`. |

### Hệ quả của mức tối thiểu ExcelApi 1.1

| Nhu cầu | API mới (không dùng trực tiếp) | Cách làm với 1.1 |
|---|---|---|
| Kiểm tra sheet tồn tại | `getItemOrNullObject` (1.4) | Load `worksheets.items/name` rồi so khớp tên. |
| Xác định ô đầu của vùng merge (R2.2) | `getMergedAreasOrNullObject` (1.13) | Người dùng chọn ô → lấy `getSelectedRange().getCell(0, 0)`. |
| Lưu cấu hình mapping (R2.8) | — | `Office.context.document.settings` (Common API). |
| Gọi mạng (kiểm tra phiên bản) | `fetch` (không có trên IE11) | `XMLHttpRequest` hoặc polyfill. |

### Phân tầng mã nguồn

```
src/
  core/        Business logic thuần TypeScript, KHÔNG phụ thuộc Office.js (unit test được)
    dates/       Chuyển đổi Excel serial date ↔ ngày, đọc text ngày
    text/        Chuẩn hóa text tiếng Việt (bỏ dấu, không dùng String.normalize vì IE11)
    address/     Tiện ích thuần về địa chỉ Excel (chữ cột, địa chỉ ô, bỏ tên sheet)
    dailyLog/    F1: nhận diện cột, kiểm tra dòng, trải công việc theo ngày, tên sheet
    mapping/     F2: mapping cột → ô, cột Ngày, danh sách ngày điều hướng, kế hoạch ghi, cấu hình lưu
    version/     F3: so sánh phiên bản, đọc/kiểm tra version.json
  excel/       Adapter Office.js: đọc/ghi range, tạo sheet, lưu settings
  commands/    Function file: lệnh của các nút Ribbon (F1, F2, F3), mở hộp thoại, điều phối excel + core
  dialog/      Hộp thoại React (chọn cột, gán cột, danh sách gán, đến ngày, xác nhận, thông báo, hướng dẫn, phiên bản)
  taskpane/    Trang chào mừng (chỉ hiện khi kích hoạt add-in)
  version/     Phiên bản đang chạy, tải version.json (XMLHttpRequest), đường dẫn tài liệu/file mẫu
```

Nguyên tắc: `commands` → gọi `excel` (I/O) và `core` (xử lý); `dialog` chỉ hiển thị/thu thập lựa chọn, giao tiếp với `commands` theo `core/dialog/protocol`; `core` không import `excel` hay `Office`.

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
