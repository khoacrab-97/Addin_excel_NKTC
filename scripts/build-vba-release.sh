#!/usr/bin/env bash
# Đóng gói bản phát hành add-in VBA "Nhật ký TC" từ 3 tệp do vba/build/build.ps1 tạo trên Windows (có Excel)
# và đã tải lên vba/release/: NhatKyThiCong.xlam, NhatKyThiCong_Core.xlam, version.txt.
#   1. Kiểm tra version.txt (phiên bản, tên tệp, SHA-256, kích thước) khớp đúng tệp lõi; hai tệp .xlam là ZIP.
#   2. Chép lõi + version.txt vào <OUT>/vba/ (kênh tự cập nhật trên GitHub Pages).
#   3. Dựng bộ cài <INSTALLER_DIR>/NhatKyTC_Setup_<phiên bản>.exe bằng makensis.
# Dùng: scripts/build-vba-release.sh <OUT> <INSTALLER_DIR>   (in ra "version=<x.y.z>" và "sha256=<bộ cài>")
set -euo pipefail

OUT=${1:?Thiếu thư mục OUT}
INSTALLER_DIR=${2:?Thiếu thư mục bộ cài}
REL=vba/release
LOADER="$REL/NhatKyThiCong.xlam"
CORE="$REL/NhatKyThiCong_Core.xlam"
MANIFEST="$REL/version.txt"

fail() { echo "LỖI: $*" >&2; exit 1; }

for f in "$LOADER" "$CORE" "$MANIFEST"; do [ -f "$f" ] || fail "thiếu $f"; done
field() { grep -E "^$1=" "$MANIFEST" | head -n 1 | cut -d= -f2- | tr -d '\r'; }
VERSION=$(field version)
FILE=$(field file)
SHA=$(field sha256)
SIZE=$(field size)
[[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || fail "version.txt: phiên bản không hợp lệ ($VERSION)"
[ "$FILE" = "NhatKyThiCong_Core.xlam" ] || fail "version.txt: file phải là NhatKyThiCong_Core.xlam"
[[ "$SHA" =~ ^[0-9a-f]{64}$ ]] || fail "version.txt: sha256 không hợp lệ"
[[ "$SIZE" =~ ^[0-9]+$ ]] || fail "version.txt: size không hợp lệ"
ACTUAL_SHA=$(sha256sum "$CORE" | cut -d' ' -f1)
ACTUAL_SIZE=$(stat -c %s "$CORE")
[ "$ACTUAL_SHA" = "$SHA" ] || fail "SHA-256 của lõi ($ACTUAL_SHA) khác version.txt — tải lại đủ 3 tệp của cùng một lần build"
[ "$ACTUAL_SIZE" = "$SIZE" ] || fail "kích thước lõi ($ACTUAL_SIZE) khác version.txt"
for f in "$LOADER" "$CORE"; do [ "$(head -c 2 "$f")" = "PK" ] || fail "$f không phải tệp .xlam (ZIP)"; done

mkdir -p "$OUT/vba" "$INSTALLER_DIR"
cp "$CORE" "$MANIFEST" "$OUT/vba/"
EXE="$(realpath "$INSTALLER_DIR")/NhatKyTC_Setup_$VERSION.exe"
makensis -V2 -DVERSION="$VERSION" -DLOADER="$(realpath "$LOADER")" -DCORE="$(realpath "$CORE")" -DOUTFILE="$EXE" installer/NhatKyTC.nsi >&2
EXE_SHA=$(sha256sum "$EXE" | cut -d' ' -f1)
echo "version=$VERSION"
echo "sha256=$EXE_SHA"
