import { describe, expect, it } from "vitest";
import { MAX_CELL_LENGTH, MAX_SPAN_DAYS, buildDailyLog } from "./buildDailyLog";
import type { WorkItem } from "./workItems";

const NOV_1 = 46327;

function item(content: string, start: number, end: number, rowNumber = 2): WorkItem {
  return { rowNumber, content, start, end };
}

describe("buildDailyLog", () => {
  it("ví dụ BUSINESS_SPEC F1.3 (R1.1, R1.2, R1.4)", () => {
    const result = buildDailyLog([
      item("Đào đất hố ga", NOV_1, NOV_1 + 2),
      item("Lắp đặt cống D600", NOV_1 + 1, NOV_1 + 1),
    ]);
    expect(result).toEqual({
      ok: true,
      firstDay: NOV_1,
      lastDay: NOV_1 + 2,
      entries: [
        { day: NOV_1, content: "- Đào đất hố ga", itemCount: 1 },
        { day: NOV_1 + 1, content: "- Đào đất hố ga\n- Lắp đặt cống D600", itemCount: 2 },
        { day: NOV_1 + 2, content: "- Đào đất hố ga", itemCount: 1 },
      ],
    });
  });

  it("R1.3: ngày không có công việc vẫn có dòng, nội dung trống", () => {
    const result = buildDailyLog([item("A", NOV_1, NOV_1), item("B", NOV_1 + 3, NOV_1 + 3)]);
    expect(result.ok && result.entries.map((entry) => entry.content)).toEqual(["- A", "", "", "- B"]);
  });

  it("R1.4: giữ thứ tự dòng input, không sắp xếp theo ngày bắt đầu", () => {
    const result = buildDailyLog([item("Sau", NOV_1 + 1, NOV_1 + 1), item("Trước", NOV_1, NOV_1 + 1)]);
    expect(result.ok && result.entries[1]?.content).toBe("- Sau\n- Trước");
  });

  it("báo lỗi khi không có công việc", () => {
    expect(buildDailyLog([])).toEqual({ ok: false, error: "Không có công việc hợp lệ để xuất." });
  });

  it("R1.10: chặn khoảng thời gian vượt giới hạn", () => {
    const atLimit = buildDailyLog([item("A", NOV_1, NOV_1 + MAX_SPAN_DAYS - 1)]);
    expect(atLimit.ok).toBe(true);
    const overLimit = buildDailyLog([item("A", NOV_1, NOV_1 + MAX_SPAN_DAYS)]);
    expect(overLimit.ok).toBe(false);
    expect(!overLimit.ok && overLimit.error).toContain("vượt giới hạn");
  });

  it("chặn nội dung một ngày vượt giới hạn ký tự của ô Excel", () => {
    const longText = "x".repeat(MAX_CELL_LENGTH);
    const result = buildDailyLog([item(longText, NOV_1, NOV_1)]);
    expect(result).toEqual({
      ok: false,
      error: `Nội dung ngày 01/11/2026 vượt quá ${MAX_CELL_LENGTH} ký tự (giới hạn một ô Excel).`,
    });
  });
});
