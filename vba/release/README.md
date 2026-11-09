# Bản dựng phát hành (add-in VBA)

Thư mục này chứa **3 tệp của cùng một lần build** trên Windows (`vba/build/build.cmd` → `vba/dist/`):

| Tệp | Vai trò |
|---|---|
| `NhatKyThiCong.xlam` | Bộ nạp: Excel mở khi khởi động, thay lõi khi có bản cập nhật đã tải, rồi mở lõi |
| `NhatKyThiCong_Core.xlam` | Lõi: tab Ribbon "Nhật ký TC", toàn bộ chức năng, tự cập nhật |
| `version.txt` | Phiên bản, tên tệp, SHA-256, kích thước của lõi (build.ps1 sinh ra) |

Khi các tệp này lên nhánh `main`, workflow *Deploy GitHub Pages*:

1. kiểm tra `version.txt` khớp đúng tệp lõi (`scripts/build-vba-release.sh`);
2. đăng lõi + `version.txt` lên `https://khoacrab-97.github.io/Addin_excel_NKTC/vba/` (kênh tự cập nhật);
3. dựng bộ cài `NhatKyTC_Setup_<phiên bản>.exe` và tạo GitHub Release `v<phiên bản>` (nếu phiên bản đó chưa phát hành).

Không sửa tay các tệp ở đây; muốn phát hành bản mới thì tăng `APP_VERSION` (vba/src/modNKTC_Ribbon.bas), build lại và tải lên đủ 3 tệp.
