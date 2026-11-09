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

    def test_strip_diacritics(self):
        self.assertEqual(strip_diacritics("Đường đi Nhật ký"), "Duong di Nhat ky")


if __name__ == "__main__":
    unittest.main()
