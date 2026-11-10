"""Chuyển mã nguồn VBA viết bằng UTF-8 (có tiếng Việt) sang ASCII để nạp vào trình soạn VBA.

Trình soạn VBA (VBE) không lưu được chữ tiếng Việt có dấu. Quy ước (giống TienBangChu):
- Chuỗi có ký tự ngoài ASCII  -> U("...\\uXXXX...") ; hàm U() trong modNKTC_Text giải mã lại.
- Chú thích (sau dấu ' ngoài chuỗi) -> bỏ dấu tiếng Việt.
- Dòng khai báo Const không được chứa chuỗi tiếng Việt (VBA không cho gọi hàm trong Const) -> báo lỗi.
- Chuỗi gốc chứa "\\u" -> báo lỗi (U() sẽ hiểu nhầm).
- Dòng kết quả dài quá 1023 ký tự (giới hạn VBA) -> báo lỗi.

Chạy: python3 vba/tools/vba_ascii.py <thư mục nguồn> <thư mục đích>
"""
import sys
import unicodedata
from pathlib import Path

MAX_LINE = 1023
# .frmcode: mã của UserForm (bố cục ở vba/forms/<tên>.json); build.ps1 tạo form rồi chép mã này vào.
# .doccode: mã của ThisWorkbook; build.ps1 chép vào module ThisWorkbook.
SOURCE_SUFFIXES = (".bas", ".cls", ".frmcode", ".doccode")


class ConvertError(Exception):
    pass


def strip_diacritics(text: str) -> str:
    text = text.replace("đ", "d").replace("Đ", "D")
    decomposed = unicodedata.normalize("NFD", text)
    ascii_text = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return "".join(ch if ord(ch) < 128 else "?" for ch in ascii_text)


def escape_literal(body: str) -> str:
    """body: nội dung giữa hai dấu nháy, đã ở dạng VBA ("" là một dấu nháy)."""
    if "\\u" in body:
        raise ConvertError('chuỗi chứa "\\u" không chuyển được sang U()')
    out = []
    for ch in unicodedata.normalize("NFC", body):
        code = ord(ch)
        if code < 128:
            out.append(ch)
        elif code <= 0xFFFF:
            out.append("\\u%04x" % code)
        else:
            raise ConvertError("ký tự ngoài BMP: %r" % ch)
    return 'U("' + "".join(out) + '")'


def convert_line(line: str) -> str:
    result = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            while True:
                if j >= n:
                    raise ConvertError("chuỗi không đóng")
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            body = line[i + 1 : j]
            if any(ord(c) > 127 for c in body):
                result.append(escape_literal(body))
            else:
                result.append(line[i : j + 1])
            i = j + 1
        elif ch == "'":
            result.append(strip_diacritics(line[i:]))
            break
        else:
            if ord(ch) > 127:
                raise ConvertError("ký tự ngoài ASCII ngoài chuỗi/chú thích: %r" % ch)
            result.append(ch)
            i += 1
    converted = "".join(result)
    code_part = converted.split("'", 1)[0].lstrip().lower()
    if 'U("' in converted and (code_part.startswith("const ") or " const " in " " + code_part):
        raise ConvertError("Const không được chứa chuỗi tiếng Việt")
    if len(converted) > MAX_LINE:
        raise ConvertError("dòng dài %d ký tự (tối đa %d)" % (len(converted), MAX_LINE))
    return converted


def convert_text(text: str, name: str = "<src>") -> str:
    lines = text.replace("\r\n", "\n").split("\n")
    out = []
    for number, line in enumerate(lines, 1):
        try:
            out.append(convert_line(line))
        except ConvertError as error:
            raise ConvertError("%s:%d: %s" % (name, number, error)) from None
    return "\r\n".join(out)


# Thư mục vba/ của repo: thư mục nguồn và đích phải nằm trong đây (không đọc/ghi ra ngoài repo).
VBA_DIR = Path(__file__).resolve().parent.parent


def inside_vba_dir(arg: str) -> Path:
    path = Path(arg).resolve()
    if not path.is_relative_to(VBA_DIR):
        raise ConvertError("%s nằm ngoài thư mục %s" % (arg, VBA_DIR))
    return path


def main(argv: list) -> int:
    if len(argv) != 3:
        print(__doc__)
        return 2
    src, dst = inside_vba_dir(argv[1]), inside_vba_dir(argv[2])
    dst.mkdir(parents=True, exist_ok=True)
    count = 0
    for path in sorted(src.iterdir()):
        if path.suffix.lower() not in SOURCE_SUFFIXES:
            continue
        converted = convert_text(path.read_text(encoding="utf-8"), path.name)
        (dst / path.name).write_bytes(converted.encode("ascii"))
        count += 1
    print("%d tệp -> %s" % (count, dst))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv))
    except ConvertError as error:
        print("LỖI: %s" % error, file=sys.stderr)
        sys.exit(1)
