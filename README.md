# ConstructionDailyLog — Add-in Excel "Nhật ký TC"

Add-in Excel (VBA, `.xlam`) lập nhật ký thi công theo ngày từ bảng tiến độ và xuất dữ liệu từng ngày vào biểu mẫu.
Tài liệu dự án: [`MASTER_CONTEXT.md`](MASTER_CONTEXT.md) · [`ROADMAP.md`](ROADMAP.md) · [`docs/BUSINESS_SPEC.md`](docs/BUSINESS_SPEC.md) · Quy tắc phát triển: [`CLAUDE.md`](CLAUDE.md).

## Cài đặt cho người dùng

[**Releases mới nhất**](https://github.com/khoacrab-97/Addin_excel_NKTC/releases/latest) → `NhatKyTC_Setup_<phiên bản>.exe` (Excel 2016 trở về sau / Microsoft 365 trên Windows, không cần quyền quản trị). Tab **Nhật ký TC** luôn hiện; add-in tự cập nhật mỗi ngày. Chi tiết: [`docs/HUONG_DAN_CAI_DAT.md`](docs/HUONG_DAN_CAI_DAT.md). Kiểm thử nghiệm thu: [`docs/KIEM_THU_NGHIEM_THU.md`](docs/KIEM_THU_NGHIEM_THU.md).

> Bản Office.js cũ (tab "Nhật ký thi công", 1.x – 2.0.1) đã gỡ khỏi repo ngày 10/11/2026 (Phase 13 trong ROADMAP); mã cũ còn trong lịch sử Git (tag `v2.0.1`).

## Chức năng (tab Nhật ký TC)

| Nhóm | Nút |
|---|---|
| Tiến độ | **Tạo bảng** (F4) · **Xuất theo ngày** (F1) |
| Form mẫu | **Chèn form mẫu** · **Xuất vào form** (F2) |
| Chuyển ngày | Ngày đầu · Ngày trước · Ngày sau · Ngày cuối |
| Tiện ích | Hướng dẫn · Phiên bản · Kiểm tra cập nhật (F3) |

## Phát triển

Mã nguồn, build (Windows + Excel), kiểm thử và quy trình phát hành: [`vba/README.md`](vba/README.md).

| Việc | Cách làm |
|---|---|
| Sửa mã | `vba/src/*.bas`, `*.frmcode`, `vba/loader/` (UTF-8, tiếng Việt có dấu) → `python3 vba/tools/vba_ascii.py vba/src vba/build/src-ascii` và `python3 vba/tools/vba_ascii.py vba/loader vba/build/loader-ascii` |
| Kiểm thử công cụ chuyển mã | `python3 -m unittest discover -s vba/tools -p "test_*.py"` |
| Build + kiểm thử phần lõi VBA | Windows có Excel: nháy đúp `vba/build/build.cmd` → `vba/dist/` (kết quả `ket-qua-build.txt`) |
| Phát hành | Tăng `APP_VERSION` → build → đưa 3 tệp `vba/dist` vào `vba/release/` → merge vào `main` |

## Biểu mẫu

`templates/Mau_NKTC.xlsx` — biểu mẫu "Tình hình thi công hàng ngày" (được đóng gói trong add-in: nút **Chèn form mẫu**). `templates/Mau_TienDo.xlsx` — danh sách công việc mẫu (input của F1). Dữ liệu thử: `test/fixtures/test_addin_xuat_nktc.xlsx`, `test/fixtures/demo_f2.xlsx` (sinh lại bằng `python3 -I scripts/fixtures/make_nktc_workbooks.py test/fixtures/test_addin_xuat_nktc.xlsx`).

## Kiểm tra bảo mật

Kết quả kiểm tra, biện pháp đã áp dụng và cấu hình GitHub: [`docs/BAO_MAT.md`](docs/BAO_MAT.md). Skill [`security-audit`](.claude/skills/security-audit/) được cài sẵn cho Claude Code; báo cáo ghi ra `~/security-audit-skill/<repo>/run-<N>` (ngoài repo).

## Workflow

| Workflow | Kích hoạt | Nội dung |
|---|---|---|
| `.github/workflows/ci.yml` | Push lên `main` / `claude/**`, pull request | Kiểm thử công cụ chuyển mã, mã ASCII khớp nguồn, dựng bộ cài thử từ `vba/release/` (artifact `vba-installer-test`, chỉ lần push) |
| `.github/workflows/deploy-pages.yml` | Push lên `main` hoặc chạy tay | GitHub Pages: trang thông báo/tải về (`site/`), biểu mẫu, kênh tự cập nhật `/vba/`; Release `vba-v<phiên bản>` kèm bộ cài |

## Cấu trúc thư mục

```
vba/                  Add-in VBA: src (lõi), loader (bộ nạp), forms, ribbon, tools, build, release (xem vba/README.md)
installer/            Bộ cài Windows (NSIS) NhatKyTC.nsi + icon
scripts/              build-vba-release.sh (CI), icons/ (vẽ icon Ribbon), fixtures/ (sinh file Excel thử)
site/                 Trang GitHub Pages (thông báo, link tải)
assets/ribbon/        Icon các nút Ribbon
templates/            Biểu mẫu Excel
test/fixtures/        File Excel dùng thử (thư mục local/ không đưa lên repo)
docs/                 Đặc tả, hướng dẫn cài đặt, kiểm thử nghiệm thu, bảo mật
```
