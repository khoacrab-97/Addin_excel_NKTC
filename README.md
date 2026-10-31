# ConstructionDailyLog — Excel Add-in Nhật ký thi công

Excel Add-in lập nhật ký thi công theo ngày và xuất theo form mẫu.
Tài liệu dự án: [`MASTER_CONTEXT.md`](MASTER_CONTEXT.md) · [`ROADMAP.md`](ROADMAP.md) · [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md) · Quy tắc phát triển: [`CLAUDE.md`](CLAUDE.md).

## Cài đặt cho người dùng

Xem [`docs/HUONG_DAN_CAI_DAT.md`](docs/HUONG_DAN_CAI_DAT.md). Bản phát hành: <https://khoacrab-97.github.io/Addin_excel_NKTC/>. Kiểm thử nghiệm thu: [`docs/KIEM_THU_NGHIEM_THU.md`](docs/KIEM_THU_NGHIEM_THU.md).

## Yêu cầu (môi trường phát triển)

- Windows, Microsoft Excel Desktop (Microsoft 365 hoặc Excel 2016 trở về sau)
- Node.js 22 LTS (22.12 trở lên)

## Chạy thử trên Excel Desktop (môi trường dev)

```bash
npm install
```

```bash
npm start
```

`npm start` sẽ:
1. Cài chứng chỉ HTTPS cho `localhost` (lần đầu Windows hỏi xác nhận — chọn **Yes**).
2. Khởi động dev server tại `https://localhost:3000`.
3. Mở Excel và sideload add-in. Nút **Nhật ký thi công** xuất hiện trên tab **Home**.

Dừng và gỡ sideload:

```bash
npm run stop
```

## Các lệnh

| Lệnh | Mục đích |
|---|---|
| `npm run check` | Chạy toàn bộ: typecheck → lint → format:check → test → build |
| `npm run typecheck` | Kiểm tra kiểu TypeScript (strict) |
| `npm run lint` / `npm run lint:fix` | ESLint (kiểm tra / tự sửa) |
| `npm run format` / `npm run format:check` | Prettier (định dạng / kiểm tra định dạng) |
| `npm test` / `npm run test:watch` | Unit test bằng Vitest (file `src/**/*.test.ts`) |
| `npm run build` | Build production vào `dist/` (manifest trỏ tới GitHub Pages) |
| `npm run build:dev` | Build development (manifest trỏ tới `localhost`) |
| `npm run validate` | Kiểm tra manifest bằng dịch vụ của Microsoft (cần Internet) |

## Biểu mẫu

`templates/Mau_NKTC.xlsx` — biểu mẫu "Tình hình thi công hàng ngày" (chuyển từ mẫu Word, A4, 2 trang). Ô đích cho từng mục: xem `docs/BUSINESS_SPEC.md` §F2.3. Cách dùng: chép sheet `Mẫu NKTC` vào file nhật ký, rồi ở tab **Form mẫu** gán các cột của sheet theo ngày vào các ô đích.

## Kiểm thử

| Loại | Lệnh | Ghi chú |
|---|---|---|
| Unit test + hồi quy | `npm test` | Gồm hồi quy F1 trên `test/fixtures/tdtc-sample.json` (dữ liệu mẫu tự tạo, mô phỏng bảng tiến độ phân cấp). |
| Smoke test F1 + F2 | `npm run smoke` (hoặc `smoke:f1`, `smoke:f2`) | Chạy bundle production trong Chromium với Excel giả lập (`scripts/smoke/fake-excel.js`). Cần Playwright + Chromium có sẵn trong môi trường (không phải dependency của dự án). |
| Dữ liệu test | `test/fixtures/test_addin_xuat_nktc.xlsx` (tiến độ thật của chủ dự án), `test/fixtures/demo_f2.xlsx` (tiến độ thật + NK_TheoNgay có cột phụ mẫu + biểu mẫu `Mẫu NKTC`) | Sinh lại: `python3 -I scripts/fixtures/make_nktc_workbooks.py test/fixtures/test_addin_xuat_nktc.xlsx` rồi `python3 -I scripts/fixtures/xlsx_to_fixture.py test/fixtures/demo_f2.xlsx test/fixtures/demo_f2.json --all`. Dữ liệu riêng tư để trong `test/fixtures/local/` (git bỏ qua). |

Kiểm thử trên Excel thật (Verified) vẫn phải chạy trên Excel Desktop Windows.

## Phát hành

| Workflow | Kích hoạt | Nội dung |
|---|---|---|
| `.github/workflows/ci.yml` | Push lên `main` / `claude/**`, pull request | `npm run check` + validate `dist/manifest.xml` bằng dịch vụ Microsoft |
| `.github/workflows/deploy-pages.yml` | Push lên `main` hoặc chạy tay (Actions → Run workflow) | Build + đưa `dist/` lên GitHub Pages |

Phiên bản phát hành lấy từ `package.json` (`version`) và tự đồng bộ vào manifest (`x.y.z.0`) và `version.json`.

## Cấu trúc thư mục

```
manifest.xml          Manifest add-in (URL dev: https://localhost:3000)
assets/               Icon add-in
src/
  core/               Business logic thuần TypeScript (không phụ thuộc Office.js)
  taskpane/           Giao diện React (task pane)
  excel/              Tầng giao tiếp Office.js
  version/            Thông tin phiên bản
  types/              Khai báo kiểu toàn cục
scripts/
  fixtures/           Chuyển file .xlsx thành fixture JSON (mô phỏng range.values)
  smoke/              Excel giả lập + smoke test chạy trên bundle production
templates/            Biểu mẫu Excel dùng thực tế (Mau_NKTC.xlsx)
test/fixtures/        Dữ liệu test (thư mục local/ không đưa lên repo)
```
