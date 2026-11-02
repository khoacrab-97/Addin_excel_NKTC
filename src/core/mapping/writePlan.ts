import { fillDatePattern } from "../dates/datePattern";
import { toDayNumber } from "../dates/excelDate";
import type { ColumnMapping, TargetCell } from "./types";

/** Một lệnh ghi giá trị vào ô đích. */
export interface CellWrite {
  target: TargetCell;
  value: string | number | boolean;
}

/**
 * Chuyển giá trị nguồn thành giá trị ghi vào ô (R2.11):
 * - Chuỗi khác rỗng được thêm dấu `'` ở đầu để Excel giữ nguyên dạng văn bản
 *   (không hiểu "- Đào đất" thành công thức, "01/11" thành ngày) mà không phải đổi định dạng ô của form.
 * - Số và true/false giữ nguyên (ngày dạng serial hiển thị theo định dạng sẵn có của ô đích).
 */
export function toCellInput(value: unknown): string | number | boolean {
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "string") return value === "" ? "" : `'${value}`;
  return "";
}

/** Thông tin để rải ngày vào ô mẫu "Ngày … tháng … năm …" (R2.15). */
export interface DateFill {
  dateColumn: number;
  /** Nội dung chữ hiện tại của ô đích (`undefined` nếu ô không phải chữ). */
  currentText: (target: TargetCell) => string | undefined;
}

/**
 * Lập danh sách ô cần ghi cho một dòng nguồn (R2.6: chỉ các ô đích đã cấu hình).
 * Cột Ngày gán vào ô đang chứa mẫu chữ "Ngày … tháng … năm …" → ghi chữ đã rải ngày/tháng/năm (R2.15);
 * ô khác nhận ngày dạng số (hiển thị theo định dạng sẵn có của ô).
 */
export function buildWritePlan(
  row: readonly unknown[],
  mappings: readonly ColumnMapping[],
  dates?: DateFill,
): CellWrite[] {
  return mappings.map((mapping) => {
    const value = row[mapping.column];
    if (dates && mapping.column === dates.dateColumn) {
      const text = dates.currentText(mapping.target);
      const day = toDayNumber(value);
      const filled = text !== undefined && day !== undefined ? fillDatePattern(text, day) : undefined;
      if (filled !== undefined) return { target: mapping.target, value: toCellInput(filled) };
    }
    return { target: mapping.target, value: toCellInput(value) };
  });
}
