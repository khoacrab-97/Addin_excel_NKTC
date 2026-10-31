import { describe, expect, it } from "vitest";
import { parseWorkItems } from "./workItems";

const COLUMNS = { content: 0, start: 1, end: 2 };
const NOV_1 = 46327;

describe("parseWorkItems", () => {
  it("đọc dòng hợp lệ (serial date và text date) với số dòng trên sheet", () => {
    const result = parseWorkItems(
      [
        ["Đào đất hố ga", NOV_1, NOV_1 + 2],
        ["  Lắp đặt cống D600  ", "02/11/2026", "02/11/2026"],
      ],
      COLUMNS,
      5,
    );
    expect(result.items).toEqual([
      { rowNumber: 5, content: "Đào đất hố ga", start: NOV_1, end: NOV_1 + 2 },
      { rowNumber: 6, content: "Lắp đặt cống D600", start: NOV_1 + 1, end: NOV_1 + 1 },
    ]);
    expect(result.issues).toEqual([]);
  });

  it("R1.5: thiếu ngày kết thúc → công việc 1 ngày", () => {
    const result = parseWorkItems([["Đổ bê tông", NOV_1, ""]], COLUMNS, 2);
    expect(result.items[0]).toMatchObject({ start: NOV_1, end: NOV_1 });
  });

  it("R1.6: báo lỗi theo dòng và bỏ qua dòng lỗi", () => {
    const result = parseWorkItems(
      [
        ["", NOV_1, NOV_1],
        ["Thiếu ngày", "", NOV_1],
        ["Sai ngày", "31/02/2026", NOV_1],
        ["Sai ngày kết thúc", NOV_1, "abc"],
        ["Ngược ngày", NOV_1 + 3, NOV_1],
        ["Hợp lệ", NOV_1, NOV_1],
      ],
      COLUMNS,
      2,
    );
    expect(result.issues).toEqual([
      { rowNumber: 2, reason: "Thiếu nội dung công việc." },
      { rowNumber: 3, reason: "Thiếu ngày bắt đầu." },
      { rowNumber: 4, reason: "Ngày bắt đầu không hợp lệ." },
      { rowNumber: 5, reason: "Ngày kết thúc không hợp lệ." },
      { rowNumber: 6, reason: "Ngày kết thúc trước ngày bắt đầu." },
    ]);
    expect(result.items.map((item) => item.rowNumber)).toEqual([7]);
  });

  it("R1.11: dòng trống ở 3 cột chính được bỏ qua, không tính là lỗi", () => {
    const result = parseWorkItems(
      [
        ["", "", "", "Ghi chú ở cột khác"],
        [null, null, null],
      ],
      COLUMNS,
      2,
    );
    expect(result).toEqual({ items: [], issues: [], blankRowCount: 2 });
  });

  it("đọc cột theo vị trí cấu hình", () => {
    const result = parseWorkItems([[NOV_1, "x", "Công việc", NOV_1]], { content: 2, start: 0, end: 3 }, 2);
    expect(result.items[0]).toMatchObject({ content: "Công việc", start: NOV_1, end: NOV_1 });
  });

  it("chuyển nội dung dạng số thành text", () => {
    const result = parseWorkItems([[123, NOV_1, NOV_1]], COLUMNS, 2);
    expect(result.items[0]?.content).toBe("123");
  });
});
