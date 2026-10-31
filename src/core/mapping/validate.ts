import { type ColumnMapping, formatTarget, targetKey } from "./types";

/**
 * Kiểm tra cấu hình mapping (R2.1): có ít nhất một mapping, cột nằm trong vùng nguồn,
 * mỗi ô đích chỉ nhận một cột. Trả về danh sách lỗi (rỗng nếu hợp lệ).
 */
export function validateMappings(mappings: readonly ColumnMapping[], columnCount: number): string[] {
  const errors: string[] = [];
  if (mappings.length === 0) errors.push("Chưa gán cột nào vào ô của form mẫu.");

  const owners = new Map<string, number>();
  for (const mapping of mappings) {
    if (!Number.isInteger(mapping.column) || mapping.column < 0 || mapping.column >= columnCount) {
      errors.push(`Cột thứ ${mapping.column + 1} không nằm trong vùng nguồn.`);
      continue;
    }
    const key = targetKey(mapping.target);
    const owner = owners.get(key);
    if (owner !== undefined && owner !== mapping.column) {
      errors.push(`Ô ${formatTarget(mapping.target)} đang được gán cho 2 cột khác nhau.`);
    }
    owners.set(key, mapping.column);
  }
  return errors;
}
