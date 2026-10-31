const DIACRITIC_GROUPS: ReadonlyArray<[RegExp, string]> = [
  [/[àáạảãâầấậẩẫăằắặẳẵ]/g, "a"],
  [/[èéẹẻẽêềếệểễ]/g, "e"],
  [/[ìíịỉĩ]/g, "i"],
  [/[òóọỏõôồốộổỗơờớợởỡ]/g, "o"],
  [/[ùúụủũưừứựửữ]/g, "u"],
  [/[ỳýỵỷỹ]/g, "y"],
  [/đ/g, "d"],
];

/**
 * Bỏ dấu tiếng Việt (chữ thường). Không dùng `String.prototype.normalize`
 * vì Internet Explorer 11 (Excel 2016/2019) không hỗ trợ.
 */
export function removeVietnameseDiacritics(text: string): string {
  let result = text.toLowerCase();
  for (const [pattern, replacement] of DIACRITIC_GROUPS) result = result.replace(pattern, replacement);
  return result;
}

/** Chuẩn hóa tiêu đề cột để so khớp: chữ thường, bỏ dấu, gộp khoảng trắng, bỏ ký tự đặc biệt. */
export function normalizeHeader(text: string): string {
  return removeVietnameseDiacritics(text)
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}
