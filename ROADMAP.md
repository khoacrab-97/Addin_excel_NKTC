# ROADMAP — ConstructionDailyLog

> Trạng thái chỉ dùng: **Not Started · In Progress · Implemented · Tested · Verified** (định nghĩa tại `CLAUDE.md` §7).
> Mã quy tắc (R1.x, R2.x) tham chiếu [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md).

## Phase 0 — Tài liệu & nền tảng

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P0.1 | Governance (`CLAUDE.md`) | Implemented |
| P0.2 | `MASTER_CONTEXT.md`, `ROADMAP.md`, `docs/BUSINESS_SPEC.md` (draft) | Implemented |
| P0.3 | Chốt câu hỏi mở Q1–Q6 trong BUSINESS_SPEC | Not Started |
| P0.4 | Khởi tạo project Office Add-in (React + TS strict, Webpack, manifest) | Not Started |
| P0.5 | Thiết lập ESLint, Prettier, Vitest, script `lint` / `typecheck` / `test` / `build` | Not Started |
| P0.6 | Sideload và chạy task pane rỗng trên Excel Desktop Windows | Not Started |

## Phase 1 — F1 Xuất nhật ký theo ngày

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P1.1 | `core/dates`: chuyển đổi Excel serial ↔ Date, parse `dd/mm/yyyy` (R1.7) | Not Started |
| P1.2 | `core/validation`: kiểm tra dòng input, báo lỗi theo dòng (R1.5, R1.6) | Not Started |
| P1.3 | `core/dailyLog`: trải công việc theo ngày, gộp nội dung (R1.1–R1.4) + unit test | Not Started |
| P1.4 | `excel`: đọc vùng input, tạo sheet mới không trùng tên, ghi kết quả (R1.8, R1.9) | Not Started |
| P1.5 | UI F1: chọn vùng, nhận diện/chỉnh cột, xem trước, xuất | Not Started |

## Phase 2 — F2 Xuất theo form mẫu

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P2.1 | `core/mapping`: mô hình mapping, kiểm tra hợp lệ (R2.1), điều hướng ngày + unit test | Not Started |
| P2.2 | `excel`: đọc vùng nguồn, đếm cột, lấy ô đang chọn, xử lý ô merge (R2.2) | Not Started |
| P2.3 | `excel`: ghi giá trị an toàn vào ô đích — kiểm tra công thức/dữ liệu sẵn có (R2.3–R2.6) | Not Started |
| P2.4 | Lưu/đọc cấu hình mapping trong workbook (R2.8) | Not Started |
| P2.5 | UI Mapping 2 khu vực Input / Output | Not Started |
| P2.6 | UI điều hướng ngày ◀ ▶, Đầu/Cuối, nhảy tới ngày; đọc lại nguồn mỗi lần (R2.7) | Not Started |
| P2.7 | Xử lý lỗi sheet/vùng không còn tồn tại (R2.9) | Not Started |

## Phase 3 — F3 Tiện ích

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P3.1 | Màn hình Hướng dẫn + file input mẫu | Not Started |
| P3.2 | Hiển thị phiên bản, kiểm tra `version.json` | Not Started |
| P3.3 | Thông báo/hướng dẫn cập nhật | Not Started |

## Phase 4 — Phát hành & nghiệm thu

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P4.1 | Hosting bản build (theo Q5) | Not Started |
| P4.2 | Kiểm thử tích hợp trên Excel Desktop Windows với dữ liệu dự án thực tế | Not Started |
| P4.3 | Tài liệu cài đặt cho người dùng | Not Started |
