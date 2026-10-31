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
