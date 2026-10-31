import type { DateSystem } from "../dates/excelDate";
import { type BuildDailyLogResult, buildDailyLog } from "./buildDailyLog";
import { type DailyLogColumns, isCompleteColumns } from "./columns";
import { type ParseWorkItemsResult, parseWorkItems } from "./workItems";

export type PrepareDailyLogResult =
  | { ok: true; parsed: ParseWorkItemsResult; built: Extract<BuildDailyLogResult, { ok: true }> }
  | { ok: false; error: string; parsed?: ParseWorkItemsResult };

/**
 * Toàn bộ xử lý F1 trên dữ liệu vùng input: dòng đầu là tiêu đề, các dòng sau là dữ liệu.
 * @param values Giá trị vùng input, gồm cả dòng tiêu đề.
 * @param firstRowNumber Số dòng trên sheet của dòng tiêu đề.
 */
export function prepareDailyLog(
  values: readonly (readonly unknown[])[],
  columns: Partial<DailyLogColumns>,
  firstRowNumber: number,
  dateSystem: DateSystem = 1900,
): PrepareDailyLogResult {
  if (values.length < 2) {
    return { ok: false, error: "Vùng input cần có dòng tiêu đề và ít nhất một dòng dữ liệu." };
  }
  if (!isCompleteColumns(columns)) {
    return { ok: false, error: "Vui lòng chọn đủ 3 cột khác nhau: Nội dung công việc, Ngày bắt đầu, Ngày kết thúc." };
  }

  const parsed = parseWorkItems(values.slice(1), columns, firstRowNumber + 1, dateSystem);
  const built = buildDailyLog(parsed.items);
  return built.ok ? { ok: true, parsed, built } : { ok: false, error: built.error, parsed };
}

/** So sánh dòng tiêu đề để phát hiện vùng input đã đổi cấu trúc giữa lúc xem trước và lúc xuất. */
export function sameHeaders(a: readonly unknown[], b: readonly unknown[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}
