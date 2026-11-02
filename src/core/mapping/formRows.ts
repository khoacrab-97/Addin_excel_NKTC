import { parseCellReference } from "../address/reference";
import { type ColumnMapping, targetKey } from "./types";

/** Một dòng gán trong hộp thoại "Xuất vào form": cột nguồn (trái) → ô trên form (phải). */
export interface FormRow {
  column?: number;
  target: string;
}

export type FormRowsResult = { ok: true; mappings: ColumnMapping[] } | { ok: false; errors: string[] };

/** Số dòng gán tối đa (khớp giới hạn cấu hình R2.14). */
export const MAX_FORM_ROWS = 500;

/**
 * Kiểm tra các dòng gán (R2.1, R2.2, R2.12): mỗi dòng có cột và một ô đích hợp lệ; ô không ghi tên sheet
 * thuộc `defaultSheet`; một ô chỉ nhận một cột. Dòng trống hoàn toàn được bỏ qua.
 */
export function buildMappingsFromRows(
  rows: readonly FormRow[],
  columnCount: number,
  defaultSheet: string,
): FormRowsResult {
  const errors: string[] = [];
  const mappings: ColumnMapping[] = [];
  const owners = new Map<string, number>();
  if (rows.length > MAX_FORM_ROWS) return { ok: false, errors: [`Tối đa ${MAX_FORM_ROWS} dòng gán.`] };

  rows.forEach((row, index) => {
    const label = `Dòng ${index + 1}`;
    const text = row.target.trim();
    if (row.column === undefined && text === "") return;
    if (row.column === undefined || !Number.isInteger(row.column) || row.column < 0 || row.column >= columnCount) {
      errors.push(`${label}: chưa chọn cột nguồn.`);
      return;
    }
    if (text === "") {
      errors.push(`${label}: chưa chọn ô trên form.`);
      return;
    }
    const reference = parseCellReference(text);
    if (!reference) {
      errors.push(`${label}: "${text}" không phải một ô hợp lệ (ví dụ Mẫu NKTC!E2).`);
      return;
    }
    const target = { sheetName: reference.sheetName ?? defaultSheet, address: reference.address };
    const key = targetKey(target);
    const owner = owners.get(key);
    if (owner !== undefined) {
      errors.push(`${label}: ô ${target.sheetName}!${target.address} đã được gán ở dòng ${owner + 1}.`);
      return;
    }
    owners.set(key, index);
    mappings.push({ column: row.column, target });
  });

  if (errors.length === 0 && mappings.length === 0) errors.push("Chưa có dòng gán nào. Bấm + Thêm để gán cột vào ô.");
  return errors.length > 0 ? { ok: false, errors } : { ok: true, mappings };
}

/** Dòng tiêu đề có ô trống (thường do ô bị gộp — merge) → không dùng được làm cơ sở gán cột. */
export function blankHeaderColumns(headers: readonly string[]): number[] {
  return headers.flatMap((header, index) => (header.trim() === "" ? [index] : []));
}
