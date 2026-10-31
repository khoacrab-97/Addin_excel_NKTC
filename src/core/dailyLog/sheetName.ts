/** Tên sheet mặc định của kết quả F1 (R1.9). */
export const DAILY_SHEET_BASE_NAME = "NK_TheoNgay";

/**
 * Tạo tên sheet chưa tồn tại: `base`, `base_2`, `base_3`… (R1.9).
 * Excel so sánh tên sheet không phân biệt hoa/thường.
 */
export function uniqueSheetName(existingNames: readonly string[], base: string = DAILY_SHEET_BASE_NAME): string {
  const taken = new Set(existingNames.map((name) => name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let suffix = 2; ; suffix++) {
    const candidate = `${base}_${suffix}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}
