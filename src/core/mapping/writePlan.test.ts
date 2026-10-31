import { describe, expect, it } from "vitest";
import { buildWritePlan, toCellInput } from "./writePlan";

describe("toCellInput (R2.11)", () => {
  it("thêm dấu ' trước chuỗi để Excel giữ nguyên văn bản", () => {
    expect(toCellInput("- Đào đất\n- Lắp cống")).toBe("'- Đào đất\n- Lắp cống");
    expect(toCellInput("=SUM(A1)")).toBe("'=SUM(A1)");
    expect(toCellInput("01/11")).toBe("'01/11");
    expect(toCellInput("Nắng")).toBe("'Nắng");
  });

  it("giữ nguyên số, true/false; chuỗi rỗng và giá trị khác → rỗng", () => {
    expect(toCellInput(46188)).toBe(46188);
    expect(toCellInput(12)).toBe(12);
    expect(toCellInput(false)).toBe(false);
    expect(toCellInput("")).toBe("");
    expect(toCellInput(null)).toBe("");
    expect(toCellInput(undefined)).toBe("");
  });
});

describe("buildWritePlan (R2.6)", () => {
  it("chỉ tạo lệnh ghi cho các ô đích đã cấu hình", () => {
    const row = [46188, "- Đào đất", "Nắng", 12];
    const plan = buildWritePlan(row, [
      { column: 0, target: { sheetName: "Form", address: "B3" } },
      { column: 1, target: { sheetName: "Form", address: "A9" } },
      { column: 3, target: { sheetName: "Form", address: "B6" } },
    ]);
    expect(plan).toEqual([
      { target: { sheetName: "Form", address: "B3" }, value: 46188 },
      { target: { sheetName: "Form", address: "A9" }, value: "'- Đào đất" },
      { target: { sheetName: "Form", address: "B6" }, value: 12 },
    ]);
  });

  it("cột thiếu trong dòng → ghi rỗng", () => {
    expect(buildWritePlan([], [{ column: 2, target: { sheetName: "Form", address: "B3" } }])).toEqual([
      { target: { sheetName: "Form", address: "B3" }, value: "" },
    ]);
  });
});
