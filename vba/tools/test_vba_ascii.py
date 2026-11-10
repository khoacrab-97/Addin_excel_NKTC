import unittest

from vba_ascii import ConvertError, convert_line, convert_text, strip_diacritics


class ConvertTests(unittest.TestCase):
    def test_ascii_line_unchanged(self):
        self.assertEqual(convert_line('    s = "abc" & x'), '    s = "abc" & x')

    def test_vietnamese_literal_to_u(self):
        self.assertEqual(convert_line('x = "Ngày"'), 'x = U("Ng\\u00e0y")')

    def test_doubled_quote_kept(self):
        self.assertEqual(convert_line('x = "Đã ""gán"""'), 'x = U("\\u0110\\u00e3 ""g\\u00e1n""")')

    def test_comment_stripped_of_diacritics(self):
        self.assertEqual(convert_line("x = 1 ' Ngày đầu tiên"), "x = 1 ' Ngay dau tien")

    def test_apostrophe_inside_string_is_not_comment(self):
        self.assertEqual(convert_line("x = \"it's\" ' ghi chú"), "x = \"it's\" ' ghi chu")

    def test_const_with_vietnamese_rejected(self):
        with self.assertRaises(ConvertError):
            convert_line('Private Const T As String = "Ngày"')

    def test_const_ascii_allowed(self):
        self.assertEqual(convert_line('Private Const T As String = "NK"'), 'Private Const T As String = "NK"')

    def test_non_ascii_code_rejected(self):
        with self.assertRaises(ConvertError):
            convert_line("Dim ngày As Long")

    def test_unclosed_string_rejected(self):
        with self.assertRaises(ConvertError):
            convert_line('x = "abc')

    def test_backslash_u_rejected(self):
        with self.assertRaises(ConvertError):
            convert_line('x = "a\\u0041 Ngày"')

    def test_decomposed_input_normalized(self):
        decomposed = "Thứ".encode("utf-8").decode("utf-8")
        import unicodedata

        self.assertEqual(convert_line('x = "%s"' % unicodedata.normalize("NFD", decomposed)), 'x = U("Th\\u1ee9")')

    def test_crlf_output_and_line_numbers_in_errors(self):
        self.assertEqual(convert_text("a\nb"), "a\r\nb")
        with self.assertRaises(ConvertError) as ctx:
            convert_text('ok\nx = "abc', "m.bas")
        self.assertIn("m.bas:2", str(ctx.exception))

    def test_main_converts_form_code_and_skips_other_files(self):
        import tempfile
        from pathlib import Path

        from vba_ascii import VBA_DIR, main

        with tempfile.TemporaryDirectory(dir=VBA_DIR) as tmp:
            src, dst = Path(tmp, "src"), Path(tmp, "dst")
            src.mkdir()
            Path(src, "frmA.frmcode").write_text('x = "Hủy"\n', encoding="utf-8")
            Path(src, "ghi-chu.txt").write_text("Ngày", encoding="utf-8")
            self.assertEqual(main(["vba_ascii.py", str(src), str(dst)]), 0)
            self.assertEqual(Path(dst, "frmA.frmcode").read_bytes(), b'x = U("H\\u1ee7y")\r\n')
            self.assertFalse(Path(dst, "ghi-chu.txt").exists())

    def test_main_rejects_directories_outside_vba(self):
        import tempfile

        from vba_ascii import VBA_DIR, main

        with tempfile.TemporaryDirectory() as outside:
            with self.assertRaises(ConvertError):
                main(["vba_ascii.py", str(VBA_DIR / "src"), outside])
            with self.assertRaises(ConvertError):
                main(["vba_ascii.py", outside, str(VBA_DIR / "build" / "src-ascii")])
            with self.assertRaises(ConvertError):
                main(["vba_ascii.py", str(VBA_DIR / "src" / ".." / ".." / "src"), str(VBA_DIR / "build" / "x")])

    def test_strip_diacritics(self):
        self.assertEqual(strip_diacritics("Đường đi Nhật ký"), "Duong di Nhat ky")


class SourceRulesTests(unittest.TestCase):
    """Quy ước của mã nguồn trong vba/src."""

    def test_no_vba_msgbox_or_inputbox(self):
        # MsgBox/InputBox của VBA mất dấu tiếng Việt (ANSI) -> dùng ShowInfo/ShowMessage (modNKTC_UI, MessageBoxW)
        # và Application.InputBox (của Excel, Unicode).
        import re
        from pathlib import Path

        pattern = re.compile(r"(?<![.\w])(MsgBox|InputBox)\b")
        offenders = []
        vba_dir = Path(__file__).resolve().parent.parent
        for path in sorted(list(vba_dir.joinpath("src").iterdir()) + list(vba_dir.joinpath("loader").iterdir())):
            for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
                code = re.sub(r'"[^"]*"', '""', line).split("'", 1)[0]  # bỏ chuỗi rồi bỏ chú thích
                if pattern.search(code):
                    offenders.append("%s:%d" % (path.name, number))
        self.assertEqual(offenders, [])


if __name__ == "__main__":
    unittest.main()
