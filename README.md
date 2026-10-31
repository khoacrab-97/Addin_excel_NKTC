# ConstructionDailyLog — Excel Add-in Nhật ký thi công

Excel Add-in lập nhật ký thi công theo ngày và xuất theo form mẫu.
Tài liệu dự án: [`MASTER_CONTEXT.md`](MASTER_CONTEXT.md) · [`ROADMAP.md`](ROADMAP.md) · [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md) · Quy tắc phát triển: [`CLAUDE.md`](CLAUDE.md).

## Yêu cầu

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

## Cấu trúc thư mục

```
manifest.xml          Manifest add-in (URL dev: https://localhost:3000)
assets/               Icon add-in
src/
  taskpane/           Giao diện React (task pane)
  excel/              Tầng giao tiếp Office.js
  version/            Thông tin phiên bản
  types/              Khai báo kiểu toàn cục
```
