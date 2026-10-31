# Bảo mật — Add-in "Nhật ký thi công"

## 1. Kết quả kiểm tra và biện pháp đã áp dụng

Kiểm tra toàn diện bằng skill `security-audit` (mức `quick`, commit `3b1004e`): **0 lỗ hổng**; 5/6 hạng mục độ phủ "covered", 1 hạng mục "blocked" (tình trạng CVE của thư viện — cần tra cứu trực tuyến bằng `npm audit`).

| Biện pháp (phiên bản 1.1.1) | Vị trí |
|---|---|
| Quyền ghi GitHub Pages / OIDC chỉ cấp cho job deploy; job build chỉ có quyền đọc | `.github/workflows/deploy-pages.yml` |
| Chỉ phát hành từ nhánh `main` (kể cả khi chạy workflow bằng tay) | `.github/workflows/deploy-pages.yml` |
| GitHub Actions cố định theo mã commit (SHA) | `.github/workflows/*.yml` |
| `npm audit --omit=dev` trong CI và trước khi phát hành — thư viện đi kèm add-in có lỗ hổng đã công bố thì dừng | `.github/workflows/*.yml` |
| F2: khóa chọn vùng nguồn / cột Ngày / gán cột khi đang chờ xác nhận ghi đè; xác nhận chỉ áp dụng cho đúng bộ ô đích đã liệt kê | `src/taskpane/components/template/` (R2.5) |
| Giới hạn cấu hình đọc từ file Excel (tên sheet, địa chỉ vùng, số tiêu đề, số ô đích…) — vượt giới hạn thì bỏ cấu hình | `src/core/mapping/config.ts` (R2.14) |
| Content-Security-Policy: task pane chỉ chạy script của add-in và Office.js | `src/taskpane/taskpane.html` |

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
2. Mục **Deployment branches and tags** → chọn **Selected branches and tags** → **Add deployment branch or tag rule** → **Ref type: Branch**, **Name pattern:** `main` → **Add rule**.

### 2.4. Quyền mặc định của GitHub Actions

1. **Settings → Actions → General**.
2. **Fork pull request workflows from outside collaborators:** chọn **Require approval for all outside collaborators**.
3. **Workflow permissions:** chọn **Read repository contents and packages permissions**; **bỏ tích** *Allow GitHub Actions to create and approve pull requests* → **Save**.

### 2.5. Kiểm tra sau khi cấu hình

- Repo → tab **Settings → Rules → Rulesets**: ruleset `Bao ve main` ở trạng thái **Active**.
- Báo Claude "đã cấu hình xong" → Claude tạo Pull Request phát hành bản 1.1.1 → anh/chị **Merge** → workflow **Deploy GitHub Pages** chạy trên `main`.

## 3. Việc còn lại

| Việc | Ghi chú |
|---|---|
| Xác nhận CSP trên Excel Desktop Windows (WebView2) | Sau khi cập nhật 1.1.1: mở add-in, dùng thử 3 tab. Nếu khung add-in trắng hoặc mất icon → báo Claude để nới policy. Excel 2016/2019 (IE11) không áp dụng CSP. |
| `npm audit` toàn bộ (gồm công cụ phát triển) | Còn cảnh báo trong công cụ `office-addin-*` của Microsoft và `minimatch` (ESLint); không nằm trong bản phát hành. Theo dõi bản cập nhật. |
| Kiểm tra lại định kỳ | Chạy skill `security-audit` (mức `standard`) khi có thay đổi lớn. |
