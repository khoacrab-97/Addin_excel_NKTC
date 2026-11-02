import { describe, expect, it } from "vitest";
import { columnOptions } from "./options";

describe("columnOptions", () => {
  it("ghép chữ cái cột theo vị trí trên sheet và tiêu đề", () => {
    expect(columnOptions(["Ngày", "", 5], 2)).toEqual([
      { index: 0, label: "Cột C – Ngày" },
      { index: 1, label: "Cột D – (không có tiêu đề)" },
      { index: 2, label: "Cột E – 5" },
    ]);
  });
});
