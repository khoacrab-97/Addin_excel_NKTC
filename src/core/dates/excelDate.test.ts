import { describe, expect, it } from "vitest";
import {
  MAX_DAY,
  MIN_DAY,
  calendarFromDay,
  dayFromCalendar,
  formatDay,
  parseDateText,
  toDayNumber,
  toWorkbookSerial,
} from "./excelDate";

// Mốc đối chiếu với Excel: =DATE(2026,11,1) → 46327.
const NOV_1_2026 = 46327;

describe("dayFromCalendar / calendarFromDay", () => {
  it("khớp Excel serial date", () => {
    expect(dayFromCalendar({ year: 2026, month: 11, day: 1 })).toBe(NOV_1_2026);
    expect(dayFromCalendar({ year: 1900, month: 3, day: 1 })).toBe(MIN_DAY);
    expect(dayFromCalendar({ year: 9999, month: 12, day: 31 })).toBe(MAX_DAY);
    expect(calendarFromDay(NOV_1_2026)).toEqual({ year: 2026, month: 11, day: 1 });
  });

  it("từ chối ngày không tồn tại", () => {
    expect(dayFromCalendar({ year: 2026, month: 2, day: 29 })).toBeUndefined();
    expect(dayFromCalendar({ year: 2026, month: 13, day: 1 })).toBeUndefined();
    expect(dayFromCalendar({ year: 2026, month: 4, day: 31 })).toBeUndefined();
    expect(dayFromCalendar({ year: 1900, month: 2, day: 28 })).toBeUndefined();
  });

  it("chấp nhận 29/02 năm nhuận", () => {
    expect(dayFromCalendar({ year: 2028, month: 2, day: 29 })).toBeDefined();
  });
});

describe("formatDay", () => {
  it("định dạng dd/mm/yyyy", () => {
    expect(formatDay(NOV_1_2026)).toBe("01/11/2026");
    expect(formatDay(NOV_1_2026 + 60)).toBe("31/12/2026");
  });
});

describe("parseDateText (R1.7)", () => {
  it.each([
    ["01/11/2026", NOV_1_2026],
    ["1/11/2026", NOV_1_2026],
    ["01-11-2026", NOV_1_2026],
    ["01.11.2026", NOV_1_2026],
    ["  01/11/2026  ", NOV_1_2026],
    ["2026-11-01", NOV_1_2026],
  ])("đọc %j", (text, expected) => {
    expect(parseDateText(text)).toBe(expected);
  });

  it.each(["31/02/2026", "11/2026", "01/11/26", "abc", "", "01/11/2026 08:00"])("từ chối %j", (text) => {
    expect(parseDateText(text)).toBeUndefined();
  });
});

describe("toDayNumber", () => {
  it("đọc serial date, bỏ phần giờ", () => {
    expect(toDayNumber(NOV_1_2026)).toBe(NOV_1_2026);
    expect(toDayNumber(NOV_1_2026 + 0.75)).toBe(NOV_1_2026);
  });

  it("đọc text ngày", () => {
    expect(toDayNumber("01/11/2026")).toBe(NOV_1_2026);
  });

  it("từ chối giá trị không phải ngày", () => {
    expect(toDayNumber(0)).toBeUndefined();
    expect(toDayNumber(-5)).toBeUndefined();
    expect(toDayNumber(Number.NaN)).toBeUndefined();
    expect(toDayNumber(MAX_DAY + 1)).toBeUndefined();
    expect(toDayNumber(true)).toBeUndefined();
    expect(toDayNumber(null)).toBeUndefined();
    expect(toDayNumber("")).toBeUndefined();
  });

  it("quy đổi hệ ngày 1904", () => {
    const serial1904 = NOV_1_2026 - 1462;
    expect(toDayNumber(serial1904, 1904)).toBe(NOV_1_2026);
    expect(toWorkbookSerial(NOV_1_2026, 1904)).toBe(serial1904);
    expect(toWorkbookSerial(NOV_1_2026)).toBe(NOV_1_2026);
  });
});
