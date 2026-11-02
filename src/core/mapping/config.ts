import { isRangeAddress, normalizeCellAddress } from "../address/address";
import { MAX_DAY } from "../dates/excelDate";
import { type ColumnMapping, type MappingConfig, mappingKey } from "./types";

/** Khóa lưu cấu hình F2 trong `Office.context.document.settings`. */
export const MAPPING_SETTINGS_KEY = "nktc.templateMapping.v1";

/**
 * Giới hạn khi đọc cấu hình từ workbook (dữ liệu do người soạn file kiểm soát) — vượt giới hạn → bỏ cấu hình.
 * Theo giới hạn của Excel: tên sheet ≤ 31 ký tự, 16.384 cột, ô ≤ 32.767 ký tự.
 */
export const CONFIG_LIMITS = {
  sheetNameLength: 31,
  columns: 16_384,
  headerLength: 32_767,
  mappings: 500,
} as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim() !== "";
}

function isSheetName(value: unknown): value is string {
  return isNonEmptyString(value) && value.length <= CONFIG_LIMITS.sheetNameLength;
}

function isHeader(value: unknown): value is string {
  return typeof value === "string" && value.length <= CONFIG_LIMITS.headerLength;
}

function isColumnIndex(value: unknown): value is number {
  return isNonNegativeInteger(value) && value < CONFIG_LIMITS.columns;
}

/**
 * Đọc cấu hình đã lưu trong workbook (dữ liệu bên ngoài → kiểm tra đầy đủ).
 * Trả về `undefined` nếu không có hoặc không hợp lệ.
 */
export function parseMappingConfig(raw: unknown): MappingConfig | undefined {
  if (!isRecord(raw) || raw.version !== 1) return undefined;
  const { source, headers, dateColumn, mappings, currentDay, confirmedKey } = raw;
  if (!isRecord(source) || !isSheetName(source.sheetName)) return undefined;
  if (typeof source.address !== "string" || !isRangeAddress(source.address)) return undefined;
  if (!Array.isArray(headers) || headers.length > CONFIG_LIMITS.columns) return undefined;
  if (!headers.every(isHeader)) return undefined;
  if (!isColumnIndex(dateColumn) || !Array.isArray(mappings) || mappings.length > CONFIG_LIMITS.mappings) {
    return undefined;
  }

  const parsedMappings: ColumnMapping[] = [];
  for (const mapping of mappings) {
    if (!isRecord(mapping) || !isColumnIndex(mapping.column) || !isRecord(mapping.target)) return undefined;
    const { sheetName, address } = mapping.target;
    if (!isSheetName(sheetName) || typeof address !== "string") return undefined;
    const normalized = normalizeCellAddress(address);
    if (normalized === undefined) return undefined;
    parsedMappings.push({ column: mapping.column, target: { sheetName, address: normalized } });
  }

  const config: MappingConfig = {
    version: 1,
    source: { sheetName: source.sheetName, address: source.address },
    headers,
    dateColumn,
    mappings: parsedMappings,
  };
  if (isNonNegativeInteger(currentDay) && currentDay <= MAX_DAY) config.currentDay = currentDay;
  // Chỉ giữ xác nhận khi khớp đúng bộ ô đích hiện tại (dữ liệu sửa tay không thể mở rộng quyền ghi).
  if (typeof confirmedKey === "string" && confirmedKey === mappingKey(parsedMappings)) {
    config.confirmedKey = confirmedKey;
  }
  return config;
}

/** Chuỗi tiêu đề để lưu/so sánh (giá trị không phải chuỗi được chuyển thành chuỗi). */
export function headerTexts(row: readonly unknown[]): string[] {
  return row.map((value) =>
    typeof value === "string" ? value : typeof value === "number" || typeof value === "boolean" ? String(value) : "",
  );
}
