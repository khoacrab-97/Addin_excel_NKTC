# Bảo mật — Add-in "Nhật ký thi công"

## 1. Kết quả kiểm tra và biện pháp đã áp dụng

| Lần | Commit / bản | Kết quả |
|---|---|---|
| run-1 | `3b1004e` (1.1.x), mức `quick` | **0 lỗ hổng**; 5/6 hạng mục "covered", 1 "blocked" (tình trạng CVE thư viện — cần tra cứu trực tuyến) |
| run-2 (03/11/2026, trước bản chính thức 2.0.0) | `8d29a95` (1.8.0), mức `quick` | **0 lỗ hổng**; 8/9 hạng mục "covered" (cấu hình workbook, dữ liệu ô → ghi ô, khung viền nháy, tin nhắn/dữ liệu hộp thoại, metadata cập nhật & trang tĩnh, bộ cài NSIS, quyền manifest/Ribbon, CI/phát hành), 1 "blocked" (tra cứu CVE ngoại tuyến — bù bằng `npm audit --omit=dev` trong CI). Critic cuối: không thiếu hạng mục. Đây là lượt kiểm tra một phần (mức `quick`), không khẳng định đã phủ hết mọi lỗi. |

| run-3 (10/11/2026, trước bản VBA 3.0.0) | `2ecee12` (VBA 3.0.0), mức `quick`, theo phạm vi `vba/`, bộ cài `NhatKyTC.nsi`, script và workflow phát hành | **0 lỗ hổng**; 9/9 hạng mục "covered" (tự cập nhật, bộ nạp thay lõi, `settings.txt`, bộ cài + khóa `OPEN` của Excel, workflow phát hành, ràng buộc bản dựng ↔ `version.txt`, cấu hình gán trong workbook nhận từ người khác, ghi ô, chọn vùng); 1 hạng mục ngoài phạm vi (phần Office.js 2.0.1). Critic cuối: không thiếu hạng mục. Lượt kiểm tra một phần, chỉ phân tích mã nguồn (không chạy VBA). Báo cáo: `~/security-audit-skill/Addin_excel_NKTC/run-3/REPORT.md`. |

Biện pháp bổ sung ở bản 2.0.0 (từ ghi chú củng cố của run-2): mã **SHA-256** của bộ cài đăng trong mỗi GitHub Release và trên trang phát hành (`NhatKyThiCong_Setup.exe.sha256`); kiểm tra số phiên bản toàn chuỗi trong workflow phát hành; `actions/checkout` không lưu token (`persist-credentials: false`).

Biện pháp của bản VBA 3.0.0 (và củng cố từ run-3):

| Biện pháp | Vị trí |
|---|---|
| Tự cập nhật qua HTTPS (TLS 1.2+); `version.txt` kiểm tra chặt (phiên bản dạng số, tên tệp không có đường dẫn, SHA-256 64 hex); tệp tải về phải khớp kích thước, chữ ký ZIP, SHA-256; chỉ nhận phiên bản lớn hơn bản đang chạy | `vba/src/modNKTC_Update.bas` |
| Bộ nạp kiểm lại SHA-256 trước khi thay lõi, sai thì xóa; giữ `.bak`, lỗi thì khôi phục | `vba/loader/modNKTC_Loader.bas` |
| Cấu hình gán đọc từ workbook được kiểm tra (tên sheet, ô đơn, ≤ 500 dòng, cột trong phạm vi); chỉ ghi vào ô của chính workbook; văn bản ghi có tiền tố `'` | `vba/src/modNKTC_Config.bas`, `vba/src/modNKTC_FormCommands.bas`, `vba/src/modNKTC_Mapping.bas` |
| Bộ cài chạy quyền người dùng, chỉ ghi `%LOCALAPPDATA%\NhatKyTC` và `HKCU`; xóa trạng thái cập nhật cũ khi cài | `installer/NhatKyTC.nsi` |
| CI chỉ phát hành từ `main`; lõi phát hành phải khớp SHA-256/kích thước trong `version.txt` **và** phiên bản phải bằng `APP_VERSION` của mã nguồn (run-3); Release dùng tag riêng `vba-v<phiên bản>` (run-3); bộ cài thử chỉ tạo từ lần chạy push, không từ Pull Request (run-3) | `scripts/build-vba-release.sh`, `.github/workflows/*.yml` |

| Biện pháp (phiên bản 1.1.1) | Vị trí |
|---|---|
| Quyền ghi GitHub Pages / OIDC chỉ cấp cho job deploy; job build chỉ có quyền đọc | `.github/workflows/deploy-pages.yml` |
| Chỉ phát hành từ nhánh `main` (kể cả khi chạy workflow bằng tay) | `.github/workflows/deploy-pages.yml` |
| GitHub Release: job `release` riêng có quyền `contents: write`, chỉ chạy trên `main`, không cài/chạy npm — chỉ tải bộ cài đã build từ job build (quyền đọc) và gọi `gh release create` | `.github/workflows/deploy-pages.yml` |
| GitHub Actions cố định theo mã commit (SHA) | `.github/workflows/*.yml` |
| `npm audit --omit=dev` trong CI và trước khi phát hành — thư viện đi kèm add-in có lỗ hổng đã công bố thì dừng | `.github/workflows/*.yml` |
| F2: xác nhận ghi đè chỉ áp dụng cho đúng bộ ô đích đã liệt kê (R2.5) | `src/commands/templateCommands.ts` |
| Giới hạn cấu hình đọc từ file Excel (tên sheet, địa chỉ vùng, số tiêu đề, số ô đích…) — vượt giới hạn thì bỏ cấu hình | `src/core/mapping/config.ts` (R2.14) |
| Content-Security-Policy: trang chào mừng, function file, hộp thoại chỉ chạy script của add-in và Office.js | `src/taskpane/taskpane.html`, `src/commands/commands.html`, `src/dialog/dialog.html` |
| Hộp thoại (1.3.0): URL hộp thoại cùng domain với add-in (yêu cầu của Dialog API); dữ liệu vào/ra được kiểm tra kiểu và giới hạn độ dài (`src/core/dialog/protocol.ts`), dữ liệu sai coi như Hủy; dữ liệu tạm trong localStorage bị xóa khi hộp thoại đóng; mỗi lần chỉ chạy một lệnh | `src/commands/dialogHost.ts`, `src/dialog/payload.ts` |
| Xác nhận ghi đè (R2.5) lưu trong workbook gắn với đúng bộ ô đích; xác nhận đọc từ file không khớp bộ ô đích bị bỏ (không mở rộng quyền ghi) | `src/core/mapping/config.ts` |
| Bộ cài (1.2.0): chạy quyền người dùng thường, chỉ ghi `%LOCALAPPDATA%\NhatKyThiCong` và `HKCU` (đăng ký manifest + mục gỡ cài đặt); không chứa mã chạy nền; bộ cài được build trong CI từ mã nguồn, không đưa file `.exe` vào repo; script build gọi `makensis` bằng đường dẫn tuyệt đối cố định, không tìm theo `PATH` (Sonar S4036) | `installer/NhatKyThiCong.nsi`, `scripts/build-installer.js` |

## 2. Cấu hình GitHub (chủ repo tự thực hiện)

Các thiết lập dưới đây nằm ngoài mã nguồn, chỉ chủ repo đổi được. Làm theo thứ tự.

### 2.1. Đặt `main` làm nhánh mặc định

1. Mở <https://github.com/khoacrab-97/Addin_excel_NKTC/settings>.
2. Mục **Default branch** → bấm biểu tượng **⇄** (Switch to another branch) → chọn **main** → **Update** → **I understand, update the default branch**.

### 2.2. Bảo vệ nhánh `main` (Ruleset)

1. **Settings → Rules → Rulesets → New ruleset → New branch ruleset**.
2. **Ruleset Name:** `Bao ve main`. **Enforcement status:** **Active**.
3. **Bypass list:** để trống.
4. **Target branches → Add target → Include default branch** (đã đặt là `main` ở bước 2.1).
5. **Rules** — tích:
   - **Restrict deletions** (cấm xóa nhánh).
   - **Block force pushes** (cấm ghi đè lịch sử).
   - **Require a pull request before merging** → **Required approvals: 0**. Repo một người: GitHub không cho tự duyệt Pull Request của chính mình, đặt 1 sẽ không merge được.
   - **Require status checks to pass** → **Add checks** → gõ `check` → chọn **check** (job CI).
6. Bấm **Create**.

Hệ quả: mọi thay đổi vào `main` (cũng là phát hành) phải đi qua **Pull Request** và CI phải xanh. Claude sẽ tạo Pull Request; anh/chị xem rồi bấm **Merge pull request** để phát hành.

### 2.3. Môi trường `github-pages` chỉ nhận `main`

1. **Settings → Environments → github-pages**.
2. Mục **Deployment branches and tags** → chọn **Selected branches and tags**.
3. Danh sách quy tắc chỉ được còn **`main`** ("1 branch and 0 tags allowed"). GitHub thường đã tự tạo quy tắc `main` khi bật Pages (thêm lại sẽ báo *"Name branch pattern already exists: main"* — bỏ qua). Nếu còn nhánh khác (ví dụ nhánh mặc định cũ `claude/…`) → **Remove**.

### 2.4. Quyền mặc định của GitHub Actions

1. **Settings → Actions → General**.
2. **Approval for running fork pull request workflows from contributors:** chọn **Require approval for all external contributors** → **Save**.
3. **Workflow permissions:** chọn **Read repository contents and packages permissions**; **bỏ tích** *Allow GitHub Actions to create and approve pull requests* → **Save**.

> **Lưu ý thứ tự:** phải làm bước 2.1 (nhánh mặc định = `main`) trước. Nếu nhánh mặc định vẫn là nhánh khác, ruleset "Include default branch" sẽ khóa nhầm nhánh đó và `main` không được bảo vệ.

### 2.5. Kiểm tra sau khi cấu hình

- Repo → tab **Settings → Rules → Rulesets**: ruleset `Bao ve main` ở trạng thái **Active**.
- Báo Claude "đã cấu hình xong" → Claude tạo Pull Request phát hành bản 1.1.1 → anh/chị **Merge** → workflow **Deploy GitHub Pages** chạy trên `main`.

## 3. Việc còn lại

| Việc | Ghi chú |
|---|---|
| Xác nhận CSP trên Excel Desktop Windows (WebView2) | Sau khi cài 1.3.0: bấm thử các nút trên tab Ribbon. Nếu hộp thoại trắng hoặc mất icon → báo Claude để nới policy. Excel 2016/2019 (IE11) không áp dụng CSP. |
| `npm audit` toàn bộ (gồm công cụ phát triển) | Còn cảnh báo trong công cụ `office-addin-*` của Microsoft và `minimatch` (ESLint); không nằm trong bản phát hành. Theo dõi bản cập nhật. |
| Bộ cài chưa ký số (code signing) | Windows SmartScreen cảnh báo khi chạy lần đầu (More info → Run anyway). Chỉ tải bộ cài từ trang phát hành chính thức và đối chiếu mã SHA-256 (từ 2.0.0). Có thể mua chứng chỉ ký mã nếu phát hành rộng. |
| Ghi chú củng cố còn mở (run-2, không phải lỗ hổng) | Xác nhận ghi đè lưu trong workbook có thể được file nhận từ người khác đặt sẵn (chỉ áp dụng cho ô đích do chính file đó định nghĩa); origin `khoacrab-97.github.io` dùng chung cho các trang Pages của tài khoản (tên miền riêng sẽ tách biệt hơn); dọn dữ liệu tạm `localStorage nktc.dialog.*` còn sót khi Excel bị đóng đột ngột; `npm ci --ignore-scripts` cho job build; ghim phiên bản NSIS. |
| Ghi chú củng cố còn mở (run-3, bản VBA — không phải lỗ hổng) | Cần build lại VBA nên để bản 3.0.x sau: kiểm tra trường `source` của cấu hình gán như ô đích; bỏ ngoại lệ `http://localhost` của `UpdateUrl` ở bản phát hành; yêu cầu `size` dương; khớp đúng giá trị `OPEN` khi đăng ký/gỡ; so phiên bản khi cài lại; thêm SHA-256 của bộ nạp vào `version.txt`; kiểm tra trước khi ghi (sheet khóa, chuỗi dài) khi chuyển ngày. Thiết lập GitHub: yêu cầu chủ repo duyệt thay đổi `vba/release/**`. Chi tiết: báo cáo run-3 mục 6. |
| Kiểm tra lại định kỳ | Chạy skill `security-audit` (mức `standard`) khi có thay đổi lớn. |
