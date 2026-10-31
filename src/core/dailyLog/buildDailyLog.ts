import { type DayNumber, formatDay } from "../dates/excelDate";
import type { WorkItem } from "./workItems";

/** Một dòng của sheet nhật ký theo ngày. */
export interface DailyEntry {
  day: DayNumber;
  /** Nội dung đã gộp (R1.4); chuỗi rỗng nếu ngày không có công việc (R1.3). */
  content: string;
  itemCount: number;
}

export type BuildDailyLogResult =
  { ok: true; entries: DailyEntry[]; firstDay: DayNumber; lastDay: DayNumber } | { ok: false; error: string };

/** Giới hạn khoảng thời gian xuất (R1.10) để chặn lỗi nhập năm (ví dụ 2062 thay vì 2026). */
export const MAX_SPAN_DAYS = 3660;
/** Số ký tự tối đa của một ô Excel. */
export const MAX_CELL_LENGTH = 32_767;

const LINE_PREFIX = "- ";
const LINE_SEPARATOR = "\n";

/** Trải công việc theo từng ngày và gộp nội dung (R1.1–R1.4, R1.10). */
export function buildDailyLog(items: readonly WorkItem[]): BuildDailyLogResult {
  if (items.length === 0) return { ok: false, error: "Không có công việc hợp lệ để xuất." };

  let firstDay = Infinity;
  let lastDay = -Infinity;
  for (const item of items) {
    firstDay = Math.min(firstDay, item.start);
    lastDay = Math.max(lastDay, item.end);
  }

  const span = lastDay - firstDay + 1;
  if (span > MAX_SPAN_DAYS) {
    return {
      ok: false,
      error:
        `Khoảng thời gian từ ${formatDay(firstDay)} đến ${formatDay(lastDay)} (${span} ngày) ` +
        `vượt giới hạn ${MAX_SPAN_DAYS} ngày. Vui lòng kiểm tra lại ngày nhập.`,
    };
  }

  const linesByDay: string[][] = [];
  for (let i = 0; i < span; i++) linesByDay.push([]);
  for (const item of items) {
    for (let day = item.start; day <= item.end; day++) {
      linesByDay[day - firstDay]?.push(LINE_PREFIX + item.content);
    }
  }

  const entries: DailyEntry[] = [];
  for (let i = 0; i < span; i++) {
    const lines = linesByDay[i] ?? [];
    const day = firstDay + i;
    const content = lines.join(LINE_SEPARATOR);
    if (content.length > MAX_CELL_LENGTH) {
      return {
        ok: false,
        error: `Nội dung ngày ${formatDay(day)} vượt quá ${MAX_CELL_LENGTH} ký tự (giới hạn một ô Excel).`,
      };
    }
    entries.push({ day, content, itemCount: lines.length });
  }

  return { ok: true, entries, firstDay, lastDay };
}
