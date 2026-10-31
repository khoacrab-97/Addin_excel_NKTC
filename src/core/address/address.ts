/** Bỏ phần tên sheet khỏi địa chỉ Excel: `'Nhật ký'!A1:C10` → `A1:C10`. */
export function stripSheetName(address: string): string {
  const separator = address.lastIndexOf("!");
  return separator === -1 ? address : address.slice(separator + 1);
}

/** Chữ cái cột từ chỉ số cột bắt đầu từ 0: 0 → A, 25 → Z, 26 → AA. */
export function columnLetter(columnIndex: number): string {
  if (!Number.isInteger(columnIndex) || columnIndex < 0)
    throw new RangeError(`Chỉ số cột không hợp lệ: ${columnIndex}`);
  let letters = "";
  let remaining = columnIndex + 1;
  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + offset) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

/** Chỉ số cột (0-based) từ chữ cái cột: A → 0, AA → 26. */
export function columnIndexFromLetter(letters: string): number {
  let index = 0;
  for (const char of letters.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64);
  return index - 1;
}

const CELL_PATTERN = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const MAX_COLUMN_INDEX = 16_383;
const MAX_ROW = 1_048_576;

/**
 * Chuẩn hóa địa chỉ một ô (bỏ `$`, viết hoa): `$b$5` → `B5`.
 * Trả về `undefined` nếu không phải địa chỉ một ô hợp lệ.
 */
export function normalizeCellAddress(address: string): string | undefined {
  const match = CELL_PATTERN.exec(address.trim());
  if (!match) return undefined;
  const letters = (match[1] ?? "").toUpperCase();
  const row = Number(match[2]);
  if (columnIndexFromLetter(letters) > MAX_COLUMN_INDEX || row < 1 || row > MAX_ROW) return undefined;
  return `${letters}${row}`;
}

const RANGE_CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`;
const RANGE_PATTERN = new RegExp(
  String.raw`^(?:${RANGE_CELL}(?::${RANGE_CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})$`,
);

/**
 * Địa chỉ một vùng liên tục không gồm tên sheet: `B2`, `A1:I201`, `A:C` (nguyên cột), `3:5` (nguyên dòng).
 * Dùng để kiểm tra địa chỉ đọc từ dữ liệu bên ngoài (cấu hình lưu trong workbook).
 */
export function isRangeAddress(address: string): boolean {
  return address.length <= 32 && RANGE_PATTERN.test(address);
}
