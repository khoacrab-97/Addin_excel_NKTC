import { type DayNumber, calendarFromDay } from "./excelDate";

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

const DAY_WORD = "ngày";
const MONTH_WORD = "tháng";
const YEAR_WORD = "năm";

/**
 * Ô đích dạng chữ "… Ngày … tháng … năm …" (ví dụ dòng ngày của biểu mẫu, "Hà Nội, ngày ... tháng ... năm ...").
 * Trả về vị trí 3 từ khóa (không phân biệt hoa/thường) hoặc `undefined` nếu không phải mẫu ngày.
 * Dùng tìm chuỗi thay cho biểu thức chính quy để tránh backtracking.
 */
function keywordPositions(text: string): [number, number, number] | undefined {
  const lower = text.toLowerCase();
  // `toLowerCase` giữ nguyên độ dài với chữ tiếng Việt dựng sẵn (NFC) → vị trí dùng được cho chuỗi gốc.
  if (lower.length !== text.length) return undefined;
  const day = lower.indexOf(DAY_WORD);
  if (day === -1) return undefined;
  const month = lower.indexOf(MONTH_WORD, day + DAY_WORD.length);
  if (month === -1) return undefined;
  const year = lower.indexOf(YEAR_WORD, month + MONTH_WORD.length);
  if (year === -1) return undefined;
  return [day, month, year];
}

/** Ô có phải mẫu "Ngày … tháng … năm …" không. */
export function isDatePattern(text: string): boolean {
  return text.length <= 200 && keywordPositions(text) !== undefined;
}

/**
 * Rải ngày vào mẫu chữ (R2.15): ngày sau chữ "Ngày", tháng sau "tháng", năm sau "năm"; giữ phần chữ đứng trước
 * "Ngày" và cách viết hoa của từ khóa. Ví dụ "Ngày ..... tháng ..... năm ......" → "Ngày 15 tháng 06 năm 2026".
 */
export function fillDatePattern(text: string, day: DayNumber): string | undefined {
  if (text.length > 200) return undefined;
  const positions = keywordPositions(text);
  if (!positions) return undefined;
  const [d, m, y] = positions;
  const date = calendarFromDay(day);
  return (
    `${text.slice(0, d + DAY_WORD.length)} ${pad2(date.day)} ` +
    `${text.slice(m, m + MONTH_WORD.length)} ${pad2(date.month)} ` +
    `${text.slice(y, y + YEAR_WORD.length)} ${date.year}`
  );
}

const WEEKDAYS = ["Chủ nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"] as const;

/** Thứ trong tuần: "Thứ Hai", …, "Chủ nhật". */
export function weekdayName(day: DayNumber): string {
  const { year, month, day: dayOfMonth } = calendarFromDay(day);
  return WEEKDAYS[new Date(Date.UTC(year, month - 1, dayOfMonth)).getUTCDay()] ?? "";
}
