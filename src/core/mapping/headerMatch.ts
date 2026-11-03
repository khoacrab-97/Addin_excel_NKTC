import type { ColumnMapping, MappingConfig, TargetCell } from "./types";

/**
 * Khớp cột theo tên tiêu đề (R2.17): cấu hình gán lưu vị trí cột trong vùng nguồn, nhưng vùng nguồn có thể bị
 * thêm / chèn / đổi chỗ cột. Mỗi cột đã gán được tìm lại theo tên tiêu đề lúc gán:
 * - Không phân biệt hoa/thường và khoảng trắng thừa; phải đúng dấu tiếng Việt (chuẩn hóa Unicode NFC để "ế"
 *   gõ dựng sẵn và tổ hợp được coi là một).
 * - Nhiều cột trùng tên → khớp theo thứ tự xuất hiện (cột "Ghi chú" thứ 2 ↔ cột "Ghi chú" thứ 2).
 */
export function normalizeHeader(text: string): string {
  return text.normalize("NFC").trim().replace(/\s+/g, " ").toLowerCase();
}

/** Khóa của từng cột: tên chuẩn hóa + thứ tự xuất hiện của tên đó (0, 1, …). */
function headerKeys(headers: readonly string[]): string[] {
  const seen = new Map<string, number>();
  return headers.map((header) => {
    const name = normalizeHeader(header);
    const count = seen.get(name) ?? 0;
    seen.set(name, count + 1);
    return `${String(count)}\u0000${name}`;
  });
}

/**
 * Vị trí mới của từng cột cũ: `result[i]` = chỉ số trong `next` của cột `previous[i]`, `undefined` nếu không còn
 * (bị xóa / đổi tên). Tiêu đề trống không khớp với cột nào.
 */
export function matchHeaders(previous: readonly string[], next: readonly string[]): (number | undefined)[] {
  const index = new Map<string, number>();
  headerKeys(next).forEach((key, column) => {
    if (!key.endsWith("\u0000")) index.set(key, column);
  });
  return headerKeys(previous).map((key) => index.get(key));
}

/** Cột đã gán không còn trong vùng nguồn. */
export interface MissingColumn {
  header: string;
  /** Ô đích của dòng gán (không có → đó là cột Ngày). */
  target?: TargetCell;
}

export type RemapResult = { ok: true; config: MappingConfig; moved: boolean } | { ok: false; missing: MissingColumn[] };

/**
 * Đặt lại vị trí cột của cấu hình đã gán theo dòng tiêu đề mới (`headers`). Mọi cột đều tìm thấy → cấu hình mới
 * (vị trí + tiêu đề mới; `moved` = có cột đổi vị trí hoặc tiêu đề khác). Có cột không tìm thấy → danh sách cột thiếu.
 */
export function remapConfig(config: MappingConfig, headers: readonly string[]): RemapResult {
  const positions = matchHeaders(config.headers, headers);
  const missing: MissingColumn[] = [];
  const headerOf = (column: number): string => config.headers[column] ?? "";

  const dateColumn = positions[config.dateColumn];
  if (dateColumn === undefined) missing.push({ header: headerOf(config.dateColumn) });
  const mappings: ColumnMapping[] = [];
  for (const mapping of config.mappings) {
    const column = positions[mapping.column];
    if (column === undefined) missing.push({ header: headerOf(mapping.column), target: mapping.target });
    else mappings.push({ ...mapping, column });
  }
  if (missing.length > 0 || dateColumn === undefined) return { ok: false, missing };

  const moved =
    dateColumn !== config.dateColumn ||
    mappings.some((mapping, index) => mapping.column !== config.mappings[index]?.column) ||
    headers.length !== config.headers.length ||
    headers.some((header, index) => header !== config.headers[index]);
  return { ok: true, config: { ...config, headers: [...headers], dateColumn, mappings }, moved };
}
