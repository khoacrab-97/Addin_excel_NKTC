import { isRangeAddress, normalizeCellAddress } from "./address";

/** Tham chiếu tới vùng/ô người dùng nhập trong hộp thoại: `Sheet!A1:C23`, `'Mẫu NKTC'!E2`, `A1:C23`. */
export interface RangeReference {
  /** `undefined` khi người dùng không ghi tên sheet. */
  sheetName?: string;
  /** Địa chỉ không gồm tên sheet, viết hoa, bỏ `$`. */
  address: string;
}

const MAX_SHEET_NAME = 31;
const INVALID_SHEET_CHARS = /[[\]:*?/\\]/;

function parseSheetName(raw: string): string | undefined {
  let name = raw.trim();
  if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
  if (name === "" || name.length > MAX_SHEET_NAME || INVALID_SHEET_CHARS.test(name)) return undefined;
  return name;
}

/** Đọc tham chiếu vùng (một vùng liên tục). `undefined` nếu không hợp lệ. */
export function parseRangeReference(text: string): RangeReference | undefined {
  const trimmed = text.trim();
  if (trimmed === "" || trimmed.length > 80) return undefined;
  const separator = trimmed.lastIndexOf("!");
  const rawAddress = (separator === -1 ? trimmed : trimmed.slice(separator + 1)).replace(/\$/g, "").toUpperCase();
  if (!isRangeAddress(rawAddress)) return undefined;
  if (separator === -1) return { address: rawAddress };
  const sheetName = parseSheetName(trimmed.slice(0, separator));
  return sheetName === undefined ? undefined : { sheetName, address: rawAddress };
}

/** Đọc tham chiếu một ô (ô đích trên form). `undefined` nếu không phải một ô hợp lệ. */
export function parseCellReference(text: string): RangeReference | undefined {
  const reference = parseRangeReference(text);
  if (!reference) return undefined;
  const cell = normalizeCellAddress(reference.address);
  if (cell === undefined) return undefined;
  return reference.sheetName === undefined ? { address: cell } : { sheetName: reference.sheetName, address: cell };
}

/** Hiển thị tham chiếu cho người dùng: `Mẫu NKTC!E2`. */
export function formatReference(sheetName: string, address: string): string {
  return `${sheetName}!${address}`;
}
