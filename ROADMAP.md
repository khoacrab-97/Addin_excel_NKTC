# ROADMAP — ConstructionDailyLog

> Trạng thái chỉ dùng: **Not Started · In Progress · Implemented · Tested · Verified** (định nghĩa tại `CLAUDE.md` §7).
> Mã quy tắc (R1.x, R2.x) tham chiếu [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md).

## Phase 0 — Tài liệu & nền tảng

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P0.1 | Governance (`CLAUDE.md`) | Implemented |
| P0.2 | `MASTER_CONTEXT.md`, `ROADMAP.md`, `docs/BUSINESS_SPEC.md` (draft) | Implemented |
| P0.3 | Chốt câu hỏi mở Q1–Q6 trong BUSINESS_SPEC | Implemented |
| P0.4 | Khởi tạo project Office Add-in (React 17 + Fluent UI v8 + TS strict, Webpack/Babel ES5, manifest ExcelApi 1.1) | Implemented |
| P0.5 | Thiết lập ESLint, Prettier, Vitest, script `lint` / `typecheck` / `test` / `build` / `check` | Tested |
| P0.6 | Sideload và chạy task pane rỗng trên Excel Desktop Windows | Not Started |

## Phase 1 — F1 Xuất nhật ký theo ngày

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P1.1 | `core/dates`: chuyển đổi Excel serial ↔ Date, parse `dd/mm/yyyy` (R1.7) | Tested |
| P1.2 | `core/dailyLog`: nhận diện cột, kiểm tra dòng input, báo lỗi theo dòng (R1.5, R1.6, R1.11, R1.12) | Tested |
| P1.3 | `core/dailyLog`: trải công việc theo ngày, gộp nội dung (R1.1–R1.4, R1.10) + unit test | Tested |
| P1.4 | `excel`: đọc vùng input, tạo sheet mới không trùng tên, ghi kết quả (R1.8, R1.9, R1.13) | Tested |
| P1.5 | UI F1: chọn vùng, nhận diện/chỉnh cột, xem trước, xuất | Tested |
| P1.6 | Kiểm thử F1 trên Excel Desktop Windows (Excel 2016 và Microsoft 365) | Not Started |

> P1.4–P1.5 **Tested** = smoke test tự động (`npm run smoke:f1`) với Excel giả lập + file tiến độ mẫu của chủ dự án. Chưa phải **Verified** cho tới khi chạy trên Excel Desktop Windows (P1.6).

## Phase 2 — F2 Xuất theo form mẫu

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P2.1 | `core/mapping`: mô hình mapping, kiểm tra hợp lệ (R2.1), điều hướng ngày + unit test | Tested |
| P2.2 | `excel`: đọc vùng nguồn, đếm cột, lấy ô đang chọn, xử lý ô merge theo cách tương thích ExcelApi 1.1 (R2.2) | Tested |
| P2.3 | `excel`: ghi giá trị an toàn vào ô đích — kiểm tra công thức/dữ liệu sẵn có (R2.3–R2.6) | Tested |
| P2.4 | Lưu/đọc cấu hình mapping trong workbook (R2.8) | Tested |
| P2.5 | UI Mapping 2 khu vực Input / Output | Tested |
| P2.6 | UI điều hướng ngày ◀ ▶, Đầu/Cuối, nhảy tới ngày; đọc lại nguồn mỗi lần (R2.7) | Tested |
| P2.7 | Xử lý lỗi sheet/vùng không còn tồn tại (R2.9) | Tested |
| P2.8 | Kiểm thử F2 trên Excel Desktop Windows với biểu mẫu NKTC thật (xác nhận R2.11: dấu `'` không hiển thị) | Not Started |
| P2.9 | Biểu mẫu Excel "Tình hình thi công hàng ngày" chuyển từ mẫu Word (`templates/Mau_NKTC.xlsx`, BUSINESS_SPEC §F2.3) | Tested |

> P2.1 **Tested** = unit test. P2.2–P2.7 **Tested** = smoke test tự động 11 bước (`npm run smoke:f2`) với Excel giả lập + `test/fixtures/demo_f2.xlsx` (dữ liệu tiến độ thật + biểu mẫu `Mẫu NKTC` chuyển từ mẫu Word). P2.9 **Tested** = smoke F2 trên biểu mẫu này + kiểm tra bản in A4 (LibreOffice → PDF, 2 trang).

## Phase 3 — F3 Tiện ích

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P3.1 | Màn hình Hướng dẫn + file input mẫu | Not Started |
| P3.2 | Hiển thị phiên bản (đã có bản cơ bản ở tab Tiện ích từ P0.4); kiểm tra `version.json` | In Progress |
| P3.3 | Thông báo/hướng dẫn cập nhật | Not Started |

## Phase 4 — Phát hành & nghiệm thu

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P4.1 | Hosting bản build (theo Q5) | Not Started |
| P4.2 | Kiểm thử tích hợp trên Excel Desktop Windows với dữ liệu dự án thực tế | Not Started |
| P4.3 | Tài liệu cài đặt cho người dùng | Not Started |
