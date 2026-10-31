import { normalizeCellAddress } from "../address/address";
import type { ColumnMapping, MappingConfig } from "./types";

/** Khóa lưu cấu hình F2 trong `Office.context.document.settings`. */
export const MAPPING_SETTINGS_KEY = "nktc.templateMapping.v1";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim() !== "";
}

/**
 * Đọc cấu hình đã lưu trong workbook (dữ liệu bên ngoài → kiểm tra đầy đủ).
 * Trả về `undefined` nếu không có hoặc không hợp lệ.
 */
export function parseMappingConfig(raw: unknown): MappingConfig | undefined {
  if (!isRecord(raw) || raw.version !== 1) return undefined;
  const { source, headers, dateColumn, mappings, currentDay } = raw;
  if (!isRecord(source) || !isNonEmptyString(source.sheetName) || !isNonEmptyString(source.address)) return undefined;
  if (!Array.isArray(headers) || !headers.every((header) => typeof header === "string")) return undefined;
  if (!isNonNegativeInteger(dateColumn) || !Array.isArray(mappings)) return undefined;

  const parsedMappings: ColumnMapping[] = [];
  for (const mapping of mappings) {
    if (!isRecord(mapping) || !isNonNegativeInteger(mapping.column) || !isRecord(mapping.target)) return undefined;
    const { sheetName, address } = mapping.target;
    if (!isNonEmptyString(sheetName) || typeof address !== "string") return undefined;
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
  if (isNonNegativeInteger(currentDay)) config.currentDay = currentDay;
  return config;
}

/** Chuỗi tiêu đề để lưu/so sánh (giá trị không phải chuỗi được chuyển thành chuỗi). */
export function headerTexts(row: readonly unknown[]): string[] {
  return row.map((value) =>
    typeof value === "string" ? value : typeof value === "number" || typeof value === "boolean" ? String(value) : "",
  );
}
