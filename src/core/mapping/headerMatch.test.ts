import { describe, expect, it } from "vitest";
import { matchHeaders, normalizeHeader, remapConfig } from "./headerMatch";
import type { MappingConfig } from "./types";

const HEADERS = ["Ngày", "Nội dung", "Thời tiết", "Ghi chú", "Ghi chú"];

function config(overrides: Partial<MappingConfig> = {}): MappingConfig {
  return {
    version: 1,
    source: { sheetName: "NK_TheoNgay", address: "A1:E10" },
    headers: HEADERS,
    dateColumn: 0,
    mappings: [
      { column: 0, target: { sheetName: "Mẫu NKTC", address: "E2" } },
      { column: 2, target: { sheetName: "Mẫu NKTC", address: "C6" } },
      { column: 4, target: { sheetName: "Mẫu NKTC", address: "C12" } },
    ],
    currentDay: 46188,
    confirmedKey: "k",
    ...overrides,
  };
}

describe("normalizeHeader", () => {
  it("bỏ qua hoa/thường và khoảng trắng thừa, giữ dấu", () => {
    expect(normalizeHeader("  THỜI   tiết ")).toBe(normalizeHeader("Thời tiết"));
    expect(normalizeHeader("Thoi tiet")).not.toBe(normalizeHeader("Thời tiết"));
  });

  it("dấu dựng sẵn và dấu tổ hợp là một", () => {
    expect(normalizeHeader("Thời tiết".normalize("NFD"))).toBe(normalizeHeader("Thời tiết"));
  });
});

describe("matchHeaders", () => {
  it("chèn cột vào giữa → các cột sau dịch sang phải", () => {
    expect(matchHeaders(HEADERS, ["Ngày", "Nội dung", "Số công nhân", "Thời tiết", "Ghi chú", "Ghi chú"])).toEqual([
      0, 1, 3, 4, 5,
    ]);
  });

  it("cột trùng tên khớp theo thứ tự xuất hiện", () => {
    expect(matchHeaders(["Ghi chú", "A", "Ghi chú"], ["Ghi chú", "Ghi chú", "A"])).toEqual([0, 2, 1]);
  });

  it("cột bị xóa / đổi tên → undefined; tiêu đề trống không khớp", () => {
    expect(matchHeaders(["Ngày", "", "Thời tiết"], ["Ngày", "", "Thời tiết (sáng)"])).toEqual([
      0,
      undefined,
      undefined,
    ]);
  });
});

describe("remapConfig", () => {
  it("tiêu đề không đổi → giữ nguyên, moved = false", () => {
    const result = remapConfig(config(), [...HEADERS]);
    expect(result).toEqual({ ok: true, config: config(), moved: false });
  });

  it("thêm cột ở cuối → vị trí giữ nguyên nhưng tiêu đề mới được lưu", () => {
    const result = remapConfig(config(), [...HEADERS, "Kiến nghị"]);
    expect(result.ok && result.moved).toBe(true);
    expect(result.ok && result.config.headers).toEqual([...HEADERS, "Kiến nghị"]);
    expect(result.ok && result.config.mappings.map((m) => m.column)).toEqual([0, 2, 4]);
  });

  it("chèn cột trước cột Ngày, đổi chỗ cột → theo tên; giữ ngày đang xem và xác nhận", () => {
    const result = remapConfig(config(), ["STT", "ghi chú", "Ngày", "Ghi Chú", "THỜI TIẾT", "Nội dung"]);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.config.dateColumn).toBe(2);
    expect(result.config.mappings.map((m) => m.column)).toEqual([2, 4, 3]);
    expect(result.config.currentDay).toBe(46188);
    expect(result.config.confirmedKey).toBe("k");
    expect(result.moved).toBe(true);
  });

  it("cột bị xóa / đổi tên → báo cột thiếu kèm ô đích; cột Ngày thiếu không có ô đích", () => {
    const result = remapConfig(config(), ["Ngày (dd/mm)", "Nội dung", "Ghi chú", "Ghi chú"]);
    expect(result).toEqual({
      ok: false,
      missing: [
        { header: "Ngày" },
        { header: "Ngày", target: { sheetName: "Mẫu NKTC", address: "E2" } },
        { header: "Thời tiết", target: { sheetName: "Mẫu NKTC", address: "C6" } },
      ],
    });
  });
});
