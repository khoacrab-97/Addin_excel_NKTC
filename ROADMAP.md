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

## Phase 7 — F4 Tạo bảng nhập tiến độ

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P7.1 | `core/inputTable`: vị trí bảng, kiểm tra ô bắt đầu/số dòng, liệt kê ô có dữ liệu (R4.2, R4.4) + unit test | Tested |
| P7.2 | `excel/inputTableWriter`: Excel Table 4 cột, STT, định dạng, sheet mới không trùng tên, không ghi đè (R4.1, R4.3, R4.4) | Tested |
| P7.3 | Nút **Tạo bảng** (nhóm Tiến độ) + hộp thoại chọn sheet / ô bắt đầu / số dòng | Tested |
| P7.4 | Kiểm thử trên Excel Desktop Windows (`docs/KIEM_THU_NGHIEM_THU.md` mục A') | Not Started |

> P7.1–P7.3 **Tested** = unit test (`layout.test.ts`, giao thức hộp thoại) + smoke F4 5 bước (sheet mới mặc định, sheet hiện tại ô B3, vùng có dữ liệu không ghi, kiểm tra nhập liệu, điền bảng → Xuất theo ngày) với Excel giả lập có Excel Table. Chưa Verified cho tới khi chạy trên Excel thật (P7.4).

## Phase 8 — Phản hồi thử nghiệm 1.4.0 (bản 1.5.0)

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P8.1 | F1: không tự thêm "- " trước nội dung (R1.4) | Tested |
| P8.2 | Ô chọn vùng có nút ⬆ (tự nhận diện, chọn lại vùng đang chọn / gõ địa chỉ) cho Xuất theo ngày và Xuất vào form; nhãn cột "Cột A – …" (R0.4) | Tested |
| P8.3 | Nút **Chèn form mẫu** (ExcelApi 1.13; Excel cũ → hướng dẫn tải file) | Tested |
| P8.4 | Hộp thoại chung **Xuất vào form** (vùng nguồn, cột ngày, dòng gán + Thêm / Delete, Gán → ngày đầu), thay 3 nút Vùng nguồn / Gán cột / Danh sách gán; tiêu đề có ô gộp → chặn | Tested |
| P8.5 | Nút Chuyển ngày khi chưa gán: chỉ hướng dẫn (không làm mờ được — R0.5) | Tested |
| P8.6 | Kiểm thử trên Excel Desktop Windows: nút ⬆ khi hộp thoại đang mở, chèn form mẫu, Xuất vào form (`docs/KIEM_THU_NGHIEM_THU.md` mục B, C) | Not Started |

> P8.1–P8.5 **Tested** = unit test (`reference`, `formRows`, giao thức, `buildDailyLog`) + smoke: F1 7 bước, F2 13 bước (ExcelApi 1.13 và 1.1), F3 10 bước, F4 5 bước — Excel giả lập có ô gộp, chèn sheet, Dialog API. Chưa Verified (P8.6): nút ⬆ giả định hộp thoại Office không chặn thao tác trên sheet (non-modal) — cần xác nhận trên Excel thật.

## Phase 9 — Phản hồi thử nghiệm 1.5.0 (bản 1.6.0)

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P9.1 | Biểu mẫu có dòng "Ngày ..... tháng ..... năm ....."; cột Ngày rải ngày/tháng/năm vào sau từng chữ (R2.15) | Tested |
| P9.2 | Bảng chuyển ngày nổi: ngày đang xem (thứ, vị trí), gõ ngày + Enter, Đầu/Trước/Sau/Cuối (R2.16); bỏ nút Đến ngày…; thứ tự nút Ribbon Đầu · Trước · Sau · Cuối | Tested |
| P9.3 | Nút ⬆: hộp thoại chính ẩn → hộp thoại nhỏ chọn trên sheet → OK → mở lại (R0.4); dòng gán mới để trống | Tested |
| P9.4 | Kiểm thử trên Excel Desktop Windows: bảng chuyển ngày (thời gian mở dài, cập nhật tại chỗ), hộp thoại nhỏ chọn ô | Not Started |

> P9.1–P9.3 **Tested** = unit test (`datePattern`, `writePlan`, giao thức) + smoke F1 7, F2 15 bước (DialogApi 1.2 cập nhật tại chỗ và 1.1 đóng/mở lại; xác nhận ghi đè khi bảng đang mở; nút Ribbon khác đóng bảng). Chưa Verified (P9.4): kích thước nút trên Ribbon do Excel tự co giãn theo độ rộng cửa sổ — add-in không ép được; nút lệnh giữ kết nối trong lúc bảng mở (cần xác nhận Excel không ngắt sau thời gian dài).

## Phase 10 — Phản hồi thử nghiệm 1.6.0 (bản 1.7.0)

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P10.1 | Sửa lỗi bấm ⬆ xong hộp thoại đóng, không làm gì được (Excel báo 12007 khi mở hộp thoại ngay sau khi đóng hộp thoại trước → tự thử lại) | Tested |
| P10.2 | Ô chọn vùng kiểu Excel (R0.4): thanh chọn tiêu đề + ô nhập + nút ⬇; địa chỉ hiện ngay khi chọn trên sheet (DocumentSelectionChanged + DialogApi 1.2); ⬇ quay lại hộp thoại ban đầu; X giữ giá trị cũ; Excel 2016/2019 lấy vùng đang chọn khi bấm ⬇ | Tested |
| P10.3 | Gán cột Ngày vào ô "Ngày ..... tháng ..... năm .........." → `Ngày 15 tháng 06 năm 2026` trong cùng ô (R2.15), kiểm tra lại toàn luồng bằng thanh chọn | Tested |
| P10.4 | Bỏ Bảng chuyển ngày (R2.16); nhóm Chuyển ngày còn 4 nút Ngày đầu · Ngày trước · Ngày sau · Ngày cuối, icon vẽ lại cùng cỡ tam giác, căn giữa | Tested |
| P10.5 | Gộp Hướng dẫn + Phiên bản thành menu **Tiện ích ▾** (manifest Menu control) | Implemented |
| P10.6 | Kiểm thử trên Excel Desktop Windows (Microsoft 365 2610): thanh chọn cập nhật địa chỉ khi chọn, 4 nút chuyển ngày cùng cỡ, menu Tiện ích (`docs/KIEM_THU_NGHIEM_THU.md` mục B, C) | Not Started |

> P10.1–P10.4 **Tested** = unit test (giao thức `pick`) + smoke: F1 7, F2 14 bước — Excel giả lập có độ trễ đóng hộp thoại (12007), sự kiện DocumentSelectionChanged, DialogApi 1.1/1.2. P10.5 chỉ kiểm tra được cấu trúc manifest (`npm run validate` nếu có mạng); hiển thị menu cần Excel thật. Chưa Verified (P10.6): cỡ nút trên Ribbon do Excel tự co giãn theo độ rộng cửa sổ; thanh chọn giả định hộp thoại Office không chặn thao tác trên sheet.

### Bản 1.7.1 (phản hồi thử nghiệm 1.7.0)

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P10.7 | Thanh chọn: bỏ hướng dẫn "Enter" — sau khi bấm chọn trên sheet, Enter thuộc về Excel (con trỏ nhảy xuống ô dưới), add-in không chặn được; xác nhận bằng nút ⬇ (Enter chỉ dùng khi đang gõ trong ô nhập) | Tested |
| P10.8 | Viền nháy quanh vùng/ô đang chọn trong lúc thanh chọn mở (R0.6): Shape nét đứt đổi màu, xóa khi bấm ⬇ / X; dọn hình còn sót | Tested |
| P10.9 | Kiểm thử trên Excel Desktop Windows: viền nháy đúng vị trí (kể cả ô gộp, vùng lớn), không làm thay đổi vùng chọn, không để lại hình sau khi đóng thanh | Not Started |

> P10.7–P10.8 **Tested** = smoke F1 7, F2 14 bước: Excel giả lập có Shape API (vị trí khung theo ô, nét đứt, không nền, đổi màu, xóa sau khi đóng thanh, dọn hình sót, không ghi ô), ExcelApi 1.1 không vẽ. Chưa Verified (P10.9): vị trí Shape theo point của Excel thật và việc vẽ Shape có làm Excel đổi vùng chọn hay không chỉ kiểm chứng được trên Excel Desktop.

### Bản 1.8.0 — gán cột theo tên tiêu đề

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P10.10 | Vùng nguồn tự mở rộng từ ô đầu tới hết dữ liệu của sheet (R2.17a) | Tested |
| P10.11 | Khớp cột đã gán và cột Ngày theo tên tiêu đề (không phân biệt hoa/thường, khoảng trắng; trùng tên theo thứ tự) khi Chuyển ngày và trong hộp thoại Xuất vào form; cột mất → chặn, báo cột + ô đích (R2.17b–d) | Tested |
| P10.12 | Kiểm thử trên Excel Desktop Windows: chèn cột giữa, thêm ngày, đổi tên tiêu đề (`docs/KIEM_THU_NGHIEM_THU.md` C11, C12) | Not Started |

> P10.10–P10.11 **Tested** = unit test `headerMatch` (9), `formRemap` (2) + smoke F2 16 bước (chèn cột giữa C + thêm ngày 01/01/2027 → Chuyển ngày ghi đúng theo tên, không gán lại; đổi tên tiêu đề → chặn, không ghi; hộp thoại báo dòng; khôi phục tên khác hoa/thường → chạy). Chưa Verified (P10.12).

## Phase 11 — Phát hành chính thức 2.0.0

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P11.1 | Kiểm tra bảo mật cuối (`security-audit` run-2, mức `quick`, commit `8d29a95`): 0 lỗ hổng, 8/9 hạng mục covered, 1 blocked (tra cứu CVE ngoại tuyến); critic cuối không thiếu hạng mục (`docs/BAO_MAT.md` mục 1) | Tested |
| P11.2 | Củng cố phát hành: SHA-256 của bộ cài trong Release và trên trang phát hành; kiểm tra phiên bản toàn chuỗi; `persist-credentials: false` | Implemented |
| P11.3 | Phiên bản 2.0.0 — bản phát hành chính thức để chia sẻ (ghi chú Release có mô tả tính năng) | Implemented |
| P11.5 | 2.0.1: hộp thoại mở bằng đường dẫn ngắn `…/dialog.html` (thanh tiêu đề do Office ghi địa chỉ — không ẩn được, chỉ rút gọn); không đọc được dữ liệu → tự mở lại bằng đường dẫn đầy đủ | Tested |
| P11.4 | Nghiệm thu của chủ dự án trên Excel Desktop Windows: phản hồi "tạm ổn" sau bản 1.7.0 (hai lỗi nhỏ đã sửa ở 1.7.1) và "OK" sau bản 1.8.0 (03/11/2026). Các mục chi tiết P10.6, P10.9, P10.12 giữ nguyên trạng thái cho tới khi có kết quả từng bước trong `docs/KIEM_THU_NGHIEM_THU.md` | Verified |

> P11.2–P11.3 chỉ kiểm chứng được khi workflow phát hành chạy trên `main` (Release v2.0.0 có file `.exe` và mã SHA-256 khớp file tải về).

---

## Phase 12 — Chuyển sang add-in VBA (`.xlam`, bản 3.0.0)

Lý do: bản Office.js cài kiểu nhà phát triển (sideload) mất tab sau khi khởi động lại Excel; bản VBA (giống TienBangChu) có tab luôn hiện, chọn vùng bằng hộp chọn vùng gốc của Excel, không cần Internet khi dùng. Bản Office.js giữ nguyên cho tới khi chủ dự án nghiệm thu bản VBA. Mã nguồn: `vba/` (xem `vba/README.md`).

| ID | Hạng mục | Trạng thái |
|---|---|---|
| P12.1 | Quy trình build trên Windows: đổi mã sang ASCII (`vba/tools/vba_ascii.py`), nạp vào Excel, biên dịch, lưu `.xlam`, gắn tab Ribbon "Nhật ký TC" | Tested |
| P12.2 | Chuyển phần lõi (chữ, ngày, F1 nhận diện cột/trải theo ngày, F2 gán cột/chuyển ngày/khớp tên cột) + bộ kiểm thử VBA: 104/104 PASS trên Excel Windows của chủ dự án (09/11/2026) | Tested |
| P12.3 | UserForm tạo khi build từ `vba/forms/*.json` + `.frmcode`; nút **Tạo bảng** (F4) và **Xuất theo ngày** (F1) với ô chọn vùng = ô nhập + nút **Chọn** (`Application.InputBox`; RefEdit/ActiveX bị Microsoft 365 chặn khi build — 09/11/2026); thông báo tiếng Việt qua `MessageBoxW`. Build PASS 119, chủ dự án chạy thử trên Excel Windows và báo OK (09/11/2026) | Verified |
| P12.4 | Nút Chèn form mẫu (sheet mẫu nằm sẵn trong add-in, không cần Internet), Xuất vào form (danh sách dòng gán + Thêm / Cập nhật / Xóa dòng; cấu hình lưu trong file bằng CustomXMLPart), 4 nút Chuyển ngày (khớp cột theo tên, hỏi trước khi ghi đè), Hướng dẫn. Build PASS 157; chủ dự án chạy thử trên Excel Windows và báo OK (09/11/2026); tiêu đề hộp thoại có dấu sửa sau đó (`SetUnicodeTitle`) | Verified |
| P12.5 | Bộ nạp `NhatKyThiCong.xlam` + tự cập nhật hằng ngày của lõi (version.txt trên GitHub Pages `/vba/`; kiểm tra kích thước, chữ ký ZIP, SHA-256; bộ nạp kiểm lại SHA-256 trước khi thay lõi) | Implemented |
| P12.6a | Bộ cài `NhatKyTC_Setup_<phiên bản>.exe` (NSIS, cho từng người dùng, đăng ký khóa OPEN của Excel; gỡ cài đặt dồn lại các khóa OPEN). Chạy thử trên Wine: cài, cài lại, gỡ khi có add-in khác | Tested |
| P12.6b | Quy trình phát hành: 3 tệp dựng trên Windows → `vba/release/` → CI kiểm tra SHA-256, đăng kênh cập nhật, dựng bộ cài, tạo Release. Bản dựng 3.0.0 (PASS 168) đã lên `vba/release/`, CI dựng bộ cài thử thành công | Tested |
| P12.6c | Hướng dẫn cài đặt bản VBA (`docs/HUONG_DAN_CAI_DAT_VBA.md`); kiểm tra bảo mật run-3 (quick, theo phạm vi VBA: 0 lỗ hổng, 9/9 hạng mục covered); củng cố phát hành: CI so `version.txt` với `APP_VERSION`, tag riêng `vba-v<phiên bản>`, bộ cài thử chỉ từ push | Tested |
| P12.6d | Phát hành chính thức 3.0.0 (merge vào `main` → kênh cập nhật Pages `/vba/`, Release `vba-v3.0.0`) | Not Started |
| P12.7 | Nghiệm thu bản VBA trên Excel Desktop Windows: bộ cài thử 3.0.0 (artifact CI) — chủ dự án cài, tab luôn hiện, các chức năng chạy OK (10/11/2026); còn kiểm tra tự cập nhật sau khi phát hành; sau đó gỡ bản Office.js theo yêu cầu chủ dự án | In Progress |
