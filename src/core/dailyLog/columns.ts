import { normalizeHeader } from "../text/normalize";

/** Vị trí (0-based, tính trong vùng input) của 3 cột chính. */
export interface DailyLogColumns {
  content: number;
  start: number;
  end: number;
}

export type ColumnKey = keyof DailyLogColumns;

interface ColumnRule {
  /** Tiêu đề (đã chuẩn hóa) khớp chính xác. */
  exact: readonly string[];
  /** Tiêu đề (đã chuẩn hóa) chứa cụm từ này. */
  contains: readonly string[];
}

const COLUMN_RULES: Readonly<Record<ColumnKey, ColumnRule>> = {
  content: {
    exact: ["noi dung cong viec", "ten cong viec", "cong viec", "noi dung"],
    contains: ["cong viec"],
  },
  start: { exact: ["ngay bat dau", "bat dau", "tu ngay"], contains: ["bat dau"] },
  end: { exact: ["ngay ket thuc", "ket thuc", "den ngay"], contains: ["ket thuc"] },
};

const DETECTION_ORDER: readonly ColumnKey[] = ["start", "end", "content"];

/**
 * Nhận diện cột theo tiêu đề (R1.12). Ưu tiên khớp chính xác, sau đó khớp "chứa".
 * Mỗi cột chỉ được gán cho một trường. Trường không nhận diện được → không có trong kết quả.
 */
export function detectColumns(headers: readonly unknown[]): Partial<DailyLogColumns> {
  const normalized = headers.map((header) => (typeof header === "string" ? normalizeHeader(header) : ""));
  const used = new Set<number>();
  const result: Partial<DailyLogColumns> = {};

  const find = (predicate: (header: string) => boolean): number | undefined => {
    const index = normalized.findIndex((header, i) => !used.has(i) && header !== "" && predicate(header));
    return index === -1 ? undefined : index;
  };

  for (const key of DETECTION_ORDER) {
    const rule = COLUMN_RULES[key];
    const index =
      find((header) => rule.exact.includes(header)) ??
      find((header) => rule.contains.some((phrase) => header.includes(phrase)));
    if (index !== undefined) {
      result[key] = index;
      used.add(index);
    }
  }
  return result;
}

/** Kiểm tra cấu hình cột đủ 3 trường và không trùng nhau. */
export function isCompleteColumns(columns: Partial<DailyLogColumns>): columns is DailyLogColumns {
  const { content, start, end } = columns;
  if (content === undefined || start === undefined || end === undefined) return false;
  return new Set([content, start, end]).size === 3;
}
