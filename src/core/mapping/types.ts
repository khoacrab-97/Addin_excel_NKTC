/** Ô đích trên form mẫu. `address` là địa chỉ một ô đã chuẩn hóa, ví dụ `B5`. */
export interface TargetCell {
  sheetName: string;
  address: string;
}

/** Gán một cột nguồn (chỉ số 0-based trong vùng nguồn) vào một ô đích (R2.1). */
export interface ColumnMapping {
  column: number;
  target: TargetCell;
}

/** Vùng nguồn của F2 (thường là sheet nhật ký theo ngày đã bổ sung cột phụ). */
export interface SourceRef {
  sheetName: string;
  /** Địa chỉ không gồm tên sheet, ví dụ `A1:G201`. */
  address: string;
}

/** Cấu hình F2 được lưu trong workbook (R2.8). */
export interface MappingConfig {
  version: 1;
  source: SourceRef;
  /** Dòng tiêu đề lúc cấu hình, dùng để phát hiện vùng nguồn đổi cấu trúc. */
  headers: string[];
  dateColumn: number;
  mappings: ColumnMapping[];
  /** Ngày đang hiển thị trên form (DayNumber), để mở lại đúng ngày. */
  currentDay?: number;
  /**
   * Bộ ô đích người dùng đã đồng ý cho add-in ghi đè (R2.5) — khóa của `mappingKey`.
   * Gán cột thay đổi → khóa khác → hỏi xác nhận lại.
   */
  confirmedKey?: string;
}

/** Khóa so sánh ô đích (tên sheet không phân biệt hoa/thường như Excel). */
export function targetKey(target: TargetCell): string {
  return `${target.sheetName.toLowerCase()}!${target.address.toUpperCase()}`;
}

/** Khóa của bộ ô đích đang cấu hình (không phụ thuộc thứ tự gán). */
export function mappingKey(mappings: readonly ColumnMapping[]): string {
  return mappings
    .map((mapping) => targetKey(mapping.target))
    .sort()
    .join("|");
}

/** Hiển thị ô đích cho người dùng: `Form!B5`. */
export function formatTarget(target: TargetCell): string {
  return `${target.sheetName}!${target.address}`;
}
