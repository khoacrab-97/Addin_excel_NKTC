import { describe, expect, it } from "vitest";
import sample from "../../../test/fixtures/tdtc-sample.json";
import { formatDay } from "../dates/excelDate";
import { detectColumns } from "./columns";
import { prepareDailyLog } from "./prepareDailyLog";

/**
 * Hồi quy F1 trên dữ liệu mô phỏng bảng tiến độ thi công thực tế:
 * tiêu đề "TÊN HẠNG MỤC/ CÔNG TÁC", dòng tổng + dòng chi tiết phân cấp bằng khoảng trắng,
 * ngày dạng Excel serial (số thực) như `range.values` trả về.
 */
describe("F1 với bảng tiến độ phân cấp (tdtc-sample)", () => {
  const values = sample.values as unknown[][];
  const columns = detectColumns(values[0] ?? []);
  const result = prepareDailyLog(values, columns, 1);

  it("nhận diện đủ 3 cột", () => {
    expect(columns).toEqual({ content: 0, start: 1, end: 2 });
  });

  it("giữ tất cả các dòng, kể cả dòng hạng mục tổng (quyết định Q7)", () => {
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.parsed.items).toHaveLength(11);
    expect(result.parsed.issues).toEqual([]);
    expect(formatDay(result.built.firstDay)).toBe("15/06/2026");
    expect(formatDay(result.built.lastDay)).toBe("07/07/2026");
    expect(result.built.entries).toHaveLength(23);
  });

  it("gộp đúng nội dung từng ngày, bỏ khoảng trắng đầu dòng, giữ thứ tự input", () => {
    if (!result.ok) throw new Error(result.error);
    const byDate = new Map(result.built.entries.map((entry) => [formatDay(entry.day), entry]));
    expect(byDate.get("15/06/2026")?.content).toBe(
      [
        "- HM - CÔNG TÁC CHUẨN BỊ",
        "- HM - CÔNG TÁC CHUẨN BỊ",
        "- Dọn dẹp mặt bằng, dựng lán trại",
        "- LẬP HỒ SƠ HOÀN CÔNG",
        "- Lập hồ sơ nghiệm thu, hoàn công",
      ].join("\n"),
    );
    expect(byDate.get("23/06/2026")?.itemCount).toBe(6);
    expect(byDate.get("02/07/2026")?.itemCount).toBe(7);
  });
});
