import { describe, expect, it } from "vitest";
import real from "../../../test/fixtures/tdtc-real.json";
import { formatDay } from "../dates/excelDate";
import { detectColumns } from "./columns";
import { prepareDailyLog } from "./prepareDailyLog";

/** Hồi quy F1 trên bảng tiến độ thật của chủ dự án (test/fixtures/test_addin_xuat_nktc.xlsx, sheet TDTC). */
describe("F1 với bảng tiến độ thật (tdtc-real)", () => {
  const values = real.values as unknown[][];
  const columns = detectColumns(values[0] ?? []);
  const result = prepareDailyLog(values, columns, 1);
  if (!result.ok) throw new Error(result.error);
  const byDate = new Map(result.built.entries.map((entry) => [formatDay(entry.day), entry]));

  it("nhận diện cột 'TÊN HẠNG MỤC/ CÔNG TÁC', 'BẮT ĐẦU', 'KẾT THÚC'", () => {
    expect(columns).toEqual({ content: 0, start: 1, end: 2 });
  });

  it("đọc đủ 22 dòng, không có lỗi; xuất 200 ngày từ 15/06/2026 đến 31/12/2026", () => {
    expect(result.parsed.items).toHaveLength(22);
    expect(result.parsed.issues).toEqual([]);
    expect(formatDay(result.built.firstDay)).toBe("15/06/2026");
    expect(formatDay(result.built.lastDay)).toBe("31/12/2026");
    expect(result.built.entries).toHaveLength(200);
  });

  it("số dòng nội dung theo ngày khớp kết quả đã kiểm tra thủ công", () => {
    expect(byDate.get("15/06/2026")?.itemCount).toBe(5);
    expect(byDate.get("01/07/2026")?.itemCount).toBe(9);
    expect(byDate.get("06/08/2026")?.itemCount).toBe(13);
    expect(byDate.get("01/11/2026")?.itemCount).toBe(9);
    expect(byDate.get("21/12/2026")?.itemCount).toBe(11);
  });
});
