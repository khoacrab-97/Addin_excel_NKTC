import { type DateSystem, type DayNumber, formatDay, toDayNumber } from "../dates/excelDate";
import { normalizeHeader } from "../text/normalize";

/** Một ngày có thể điều hướng tới và dòng dữ liệu tương ứng trong vùng nguồn. */
export interface SourceDay {
  day: DayNumber;
  /** Chỉ số dòng trong mảng `values` của vùng nguồn (0 là dòng tiêu đề). */
  valueRowIndex: number;
  /** Số dòng trên sheet (1-based). */
  rowNumber: number;
}

export interface DayIssue {
  rowNumber: number;
  reason: string;
}

export interface DayIndex {
  /** Các ngày tăng dần, không trùng (R2.10). */
  days: SourceDay[];
  issues: DayIssue[];
}

function isBlank(value: unknown): boolean {
  return value === undefined || value === null || (typeof value === "string" && value.trim() === "");
}

/**
 * Nhận diện cột Ngày: ưu tiên tiêu đề "Ngày" (không phân biệt dấu/hoa thường), sau đó tiêu đề chứa "ngay",
 * cuối cùng là cột đầu tiên có toàn bộ ô không trống là ngày hợp lệ.
 */
export function detectDateColumn(values: readonly (readonly unknown[])[]): number | undefined {
  const headers = (values[0] ?? []).map((header) => (typeof header === "string" ? normalizeHeader(header) : ""));
  const exact = headers.indexOf("ngay");
  if (exact !== -1) return exact;
  const partial = headers.findIndex((header) => /(^| )ngay( |$)/.test(header));
  if (partial !== -1) return partial;

  const rows = values.slice(1);
  for (let column = 0; column < headers.length; column++) {
    const cells = rows.map((row) => row[column]).filter((value) => !isBlank(value));
    if (cells.length > 0 && cells.every((value) => toDayNumber(value) !== undefined)) return column;
  }
  return undefined;
}

/**
 * Lập danh sách ngày điều hướng từ vùng nguồn (dòng 0 là tiêu đề) — R2.10:
 * dòng có ô Ngày trống được bỏ qua; ngày không hợp lệ hoặc trùng được báo lỗi (giữ dòng xuất hiện trước).
 */
export function buildDayIndex(
  values: readonly (readonly unknown[])[],
  dateColumn: number,
  firstRowNumber: number,
  dateSystem: DateSystem = 1900,
): DayIndex {
  const issues: DayIssue[] = [];
  const seen = new Map<DayNumber, number>();
  const days: SourceDay[] = [];

  for (let index = 1; index < values.length; index++) {
    const rowNumber = firstRowNumber + index;
    const raw = values[index]?.[dateColumn];
    if (isBlank(raw)) continue;
    const day = toDayNumber(raw, dateSystem);
    if (day === undefined) {
      issues.push({ rowNumber, reason: "Ngày không hợp lệ." });
      continue;
    }
    const firstRow = seen.get(day);
    if (firstRow !== undefined) {
      issues.push({ rowNumber, reason: `Trùng ngày ${formatDay(day)} với dòng ${firstRow}; bỏ qua dòng này.` });
      continue;
    }
    seen.set(day, rowNumber);
    days.push({ day, valueRowIndex: index, rowNumber });
  }

  days.sort((a, b) => a.day - b.day);
  return { days, issues };
}

/** Vị trí của ngày trong danh sách; `-1` nếu không có. */
export function indexOfDay(days: readonly SourceDay[], day: DayNumber): number {
  return days.findIndex((entry) => entry.day === day);
}

/** Vị trí sau khi di chuyển `delta` bước, giới hạn trong phạm vi danh sách (R2 — điều hướng). */
export function stepIndex(length: number, current: number, delta: number): number {
  if (length === 0) return -1;
  return Math.min(length - 1, Math.max(0, current + delta));
}
