import { MAPPING_SETTINGS_KEY, parseMappingConfig } from "../core/mapping/config";
import type { MappingConfig } from "../core/mapping/types";

/** Đọc cấu hình F2 đã lưu trong workbook (R2.8). Trả về `undefined` nếu chưa có hoặc không hợp lệ. */
export function loadMappingConfig(): MappingConfig | undefined {
  const raw: unknown = Office.context.document.settings.get(MAPPING_SETTINGS_KEY);
  return parseMappingConfig(raw);
}

/** Lưu cấu hình F2 vào workbook. Cấu hình được giữ khi người dùng lưu file Excel. */
export function saveMappingConfig(config: MappingConfig): Promise<void> {
  const settings = Office.context.document.settings;
  settings.set(MAPPING_SETTINGS_KEY, config);
  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}
