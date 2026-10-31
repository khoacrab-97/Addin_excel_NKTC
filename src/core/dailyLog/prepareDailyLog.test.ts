import { describe, expect, it } from "vitest";
import { prepareDailyLog, sameHeaders } from "./prepareDailyLog";

const NOV_1 = 46327;
const HEADER = ["Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"];
const COLUMNS = { content: 0, start: 1, end: 2 };

describe("prepareDailyLog", () => {
  it("xử lý trọn vẹn: bỏ dòng tiêu đề, đánh số dòng đúng, trải theo ngày", () => {
    const result = prepareDailyLog(
      [HEADER, ["Đào đất", NOV_1, NOV_1 + 1], ["", "", "x"]],
      COLUMNS,
      3, // tiêu đề ở dòng 3 → dữ liệu bắt đầu dòng 4
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.parsed.items[0]?.rowNumber).toBe(4);
    expect(result.parsed.issues).toEqual([{ rowNumber: 5, reason: "Thiếu nội dung công việc." }]);
    expect(result.built.entries).toHaveLength(2);
  });

  it("báo lỗi khi chỉ có dòng tiêu đề", () => {
    expect(prepareDailyLog([HEADER], COLUMNS, 1)).toMatchObject({ ok: false });
  });

  it("báo lỗi khi chưa chọn đủ cột", () => {
    const result = prepareDailyLog([HEADER, ["A", NOV_1, NOV_1]], { content: 0, start: 1 }, 1);
    expect(result).toMatchObject({ ok: false });
  });

  it("trả kèm kết quả đọc dòng khi không có công việc hợp lệ", () => {
    const result = prepareDailyLog([HEADER, ["A", "sai", NOV_1]], COLUMNS, 1);
    expect(result.ok).toBe(false);
    expect(!result.ok && result.parsed?.issues).toHaveLength(1);
  });
});

describe("sameHeaders", () => {
  it("so sánh theo thứ tự và độ dài", () => {
    expect(sameHeaders(HEADER, [...HEADER])).toBe(true);
    expect(sameHeaders(HEADER, HEADER.slice(0, 2))).toBe(false);
    expect(sameHeaders(HEADER, ["x", ...HEADER.slice(1)])).toBe(false);
  });
});
