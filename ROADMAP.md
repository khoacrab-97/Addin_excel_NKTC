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
| P0.6 | Sideload và chạy add-in trên Excel Desktop Windows | Verified |

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
| P3.1 | Màn hình Hướng dẫn + file input mẫu (`templates/Mau_TienDo.xlsx`, `templates/Mau_NKTC.xlsx` phát hành kèm add-in) | Tested |
| P3.2 | Hiển thị phiên bản; kiểm tra `version.json` (R3.1–R3.3) | Tested |
| P3.3 | Thông báo bản mới đầu khung + hướng dẫn cập nhật | Tested |

> P3.x **Tested** = unit test `core/version` + smoke test tự động 7 bước (`npm run smoke:f3`): bản mới nhất, có bản mới (thông báo), bản phát triển, mất mạng / HTTP 404 / nội dung hỏng.

## Phase 4 — Phát hành & nghiệm thu

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P4.1 | Hosting bản build trên GitHub Pages (workflow `deploy-pages.yml`, tên file có mã băm, `version.json`, trang giới thiệu) | Tested |
| P4.2 | Kiểm thử tích hợp trên Excel Desktop Windows với dữ liệu dự án thực tế (`docs/KIEM_THU_NGHIEM_THU.md`) | Not Started |
| P4.3 | Hướng dẫn cài đặt: bộ cài + cách thay thế Shared Folder Catalog (`docs/HUONG_DAN_CAI_DAT.md`) | Implemented |
| P4.4 | CI trên GitHub Actions: `npm run check` + validate manifest bằng dịch vụ Microsoft (`ci.yml`) | Tested |
| P4.5 | Bộ cài Windows `NhatKyThiCong_Setup.exe` (NSIS, không cần quyền quản trị, đăng ký `HKCU\…\Wef\Developer`, gỡ qua Settings → Apps; 1.2.1: workbook kích hoạt + lối tắt Start menu); link chia sẻ = trang phát hành (GitHub Pages) hoặc GitHub Releases `v<phiên bản>` tự tạo khi merge vào `main` | Implemented |
| P4.6 | Kiểm thử bộ cài trên Windows thật: cài, tab Ribbon xuất hiện, cài đè, gỡ cài đặt (Microsoft 365 và Excel 2016) | In Progress |

## Phase 5 — Bảo mật

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P5.1 | Kiểm tra bảo mật toàn diện (skill `security-audit`, mức quick, commit `3b1004e`): 0 phát hiện, 5/6 hạng mục covered, 1 blocked (CVE thư viện — ngoại tuyến) | Implemented |
| P5.2 | Deploy: quyền `pages`/`id-token` chỉ ở job deploy; chỉ phát hành từ `main`; action cố định theo SHA; `npm audit --omit=dev` trong CI và deploy | Implemented |
| P5.3 | F2: xác nhận ghi đè gắn với đúng bộ ô đích, lưu trong workbook, không mở rộng khi đọc từ file (R2.5) | Tested |
| P5.4 | Giới hạn cấu hình đọc từ workbook (R2.14) | Tested |
| P5.5 | Content-Security-Policy cho trang chào mừng, function file, hộp thoại (`script-src 'self' https://appsforoffice.microsoft.com`) | Tested |
| P5.6 | Cấu hình GitHub: ruleset bảo vệ `main`, môi trường `github-pages` chỉ nhận `main`, nhánh mặc định `main`, quyền token mặc định chỉ đọc (`docs/BAO_MAT.md` §2) | Implemented |
| P5.7 | Xác nhận CSP không chặn Office.js / Fluent UI trên Excel Desktop Windows (WebView2) | Not Started |

> P5.3–P5.5 **Tested** = unit test (165) + smoke F1/F2/F3 có bắt vi phạm CSP qua console; CSP được kiểm chứng chặn script inline chèn vào. P5.2 chỉ có hiệu lực khi workflow chạy trên `main`. P5.6 do chủ repo cấu hình ngày 31/10/2026; đã đối chiếu qua GitHub API: nhánh mặc định `main`, ruleset trên `main` gồm deletion, non_fast_forward, pull_request (0 approval), required_status_checks (`check`).

> P4.5 **Implemented** = bộ cài build được bằng `makensis` (Linux/CI); đã giải nén kiểm tra nội dung (manifest trùng `dist/manifest.xml`). Chưa chạy trên Windows → chưa Tested/Verified (P4.6). Thử lần 1 (1.2.0, chủ dự án): cài xong không thấy add-in — bộ cài chỉ ghi registry, chưa kích hoạt; sửa ở 1.2.1. Thử lần 2 (1.2.1, chủ dự án, 02/11/2026): add-in nạp được và hiện khung bên phải trên Excel Windows (→ P0.6 Verified); chủ dự án yêu cầu chuyển chức năng lên Ribbon (Phase 6).

> P4.1 **Tested** = workflow deploy chạy thành công trên GitHub Actions (build, validate manifest, đưa lên Pages). P4.4 **Tested** = CI xanh trên GitHub ("The manifest is valid.").

## Phase 6 — Giao diện Ribbon (thay khung bên phải)

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P6.1 | Tab Ribbon riêng "Nhật ký thi công": 4 nhóm, 11 nút `ExecuteFunction`, icon 16/32/80 (BUSINESS_SPEC R0.1) | Tested |
| P6.2 | Function file `commands.html`: lệnh F1/F2/F3, một lệnh tại một thời điểm, báo lỗi bằng hộp thoại | Tested |
| P6.3 | Hộp thoại Office Dialog API 1.1: chọn cột F1, kết quả, gán cột, danh sách gán, đến ngày, xác nhận ghi đè, hướng dẫn, phiên bản; giao thức kiểm tra dữ liệu, chỉ nhận tin từ origin của add-in (R0.2) | Tested |
| P6.4 | Điều hướng ngày không mở khung: Ngày đầu/trước/sau/cuối, Đến ngày; ngày đang xem lưu trong cấu hình | Tested |
| P6.5 | Bỏ khung làm việc bên phải; trang chào mừng chỉ hiện khi kích hoạt (R0.3) | Tested |
| P6.6 | Kiểm thử tab Ribbon + hộp thoại trên Excel Desktop Windows (Microsoft 365, Excel 2016) theo `docs/KIEM_THU_NGHIEM_THU.md` | Not Started |

> P6.1–P6.5 **Tested** = unit test (`core/dialog/protocol`, `core/mapping/navigation`, `confirmedKey`) + smoke test: F1 5 bước, F2 14 bước, F3 10 bước — bấm nút bằng cách gọi hàm của function file, thao tác hộp thoại ở trang riêng với Office/Excel/Dialog API giả lập. Manifest được kiểm tra bằng dịch vụ Microsoft trong CI. Chưa Verified cho tới khi chạy trên Excel thật (P6.6).
