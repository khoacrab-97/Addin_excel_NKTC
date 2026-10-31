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

/** Lập danh sách ô cần ghi cho một dòng nguồn (R2.6: chỉ các ô đích đã cấu hình). */
export function buildWritePlan(row: readonly unknown[], mappings: readonly ColumnMapping[]): CellWrite[] {
  return mappings.map((mapping) => ({ target: mapping.target, value: toCellInput(row[mapping.column]) }));
}
