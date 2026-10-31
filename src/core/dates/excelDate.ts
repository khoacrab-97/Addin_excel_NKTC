/**
 * Ngày dạng số nguyên theo hệ ngày 1900 của Excel (Excel serial date), ví dụ 46327 = 01/11/2026.
 * Dùng số nguyên giúp tính khoảng ngày đơn giản và không bị ảnh hưởng bởi múi giờ.
 */
export type DayNumber = number;

/** Hệ ngày của workbook: 1900 (mặc định Windows) hoặc 1904. */
export type DateSystem = 1900 | 1904;

export interface CalendarDate {
  year: number;
  month: number;
  day: number;
}

const MS_PER_DAY = 86_400_000;
/** Excel serial 0 (hệ 1900, đã tính lỗi năm nhuận 1900) tương ứng 30/12/1899. */
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30);
/** Chênh lệch serial giữa hệ ngày 1900 và 1904. */
const DATE_SYSTEM_1904_OFFSET = 1462;

/** 01/03/1900: ngày nhỏ nhất được hỗ trợ (tránh vùng lỗi năm nhuận 1900 của Excel). */
export const MIN_DAY: DayNumber = 61;
/** 31/12/9999: ngày lớn nhất Excel hỗ trợ. */
export const MAX_DAY: DayNumber = 2_958_465;

function isValidDay(day: number): day is DayNumber {
  return Number.isInteger(day) && day >= MIN_DAY && day <= MAX_DAY;
}

/** Chuyển ngày/tháng/năm sang DayNumber; trả về `undefined` nếu ngày không tồn tại (ví dụ 31/02). */
export function dayFromCalendar({ year, month, day }: CalendarDate): DayNumber | undefined {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return undefined;
  if (month < 1 || month > 12 || day < 1 || day > 31) return undefined;
  const utc = Date.UTC(year, month - 1, day);
  const check = new Date(utc);
  // Date.UTC tự "tràn" ngày không hợp lệ (31/02 → 03/03) → so lại để phát hiện.
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) {
    return undefined;
  }
  const serial = Math.round((utc - EPOCH_1900_UTC) / MS_PER_DAY);
  return isValidDay(serial) ? serial : undefined;
}

/** Chuyển DayNumber sang ngày/tháng/năm. */
export function calendarFromDay(day: DayNumber): CalendarDate {
  const date = new Date(EPOCH_1900_UTC + day * MS_PER_DAY);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** Định dạng `dd/mm/yyyy`. */
export function formatDay(day: DayNumber): string {
  const { year, month, day: dayOfMonth } = calendarFromDay(day);
  return `${pad2(dayOfMonth)}/${pad2(month)}/${year}`;
}

const DMY_PATTERN = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/;
const ISO_PATTERN = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;

/** Đọc ngày dạng text: `dd/mm/yyyy` (dấu phân cách `/`, `-`, `.`) hoặc `yyyy-mm-dd`. */
export function parseDateText(text: string): DayNumber | undefined {
  const trimmed = text.trim();
  const dmy = DMY_PATTERN.exec(trimmed);
  if (dmy) return dayFromCalendar({ day: Number(dmy[1]), month: Number(dmy[2]), year: Number(dmy[3]) });
  const iso = ISO_PATTERN.exec(trimmed);
  if (iso) return dayFromCalendar({ year: Number(iso[1]), month: Number(iso[2]), day: Number(iso[3]) });
  return undefined;
}

/**
 * Chuyển giá trị ô Excel (`range.values`) sang DayNumber (R1.7).
 * - Số: Excel serial date; phần giờ (thập phân) bị bỏ.
 * - Chuỗi: xem `parseDateText`.
 * Trả về `undefined` nếu không phải ngày hợp lệ.
 */
export function toDayNumber(value: unknown, dateSystem: DateSystem = 1900): DayNumber | undefined {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return undefined;
    const day = Math.floor(value) + (dateSystem === 1904 ? DATE_SYSTEM_1904_OFFSET : 0);
    return isValidDay(day) ? day : undefined;
  }
  if (typeof value === "string") return parseDateText(value);
  return undefined;
}

/** Chuyển DayNumber sang serial theo hệ ngày của workbook để ghi vào ô. */
export function toWorkbookSerial(day: DayNumber, dateSystem: DateSystem = 1900): number {
  return dateSystem === 1904 ? day - DATE_SYSTEM_1904_OFFSET : day;
}
