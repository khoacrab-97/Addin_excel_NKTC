import { describe, expect, it } from "vitest";
import { validateMappings } from "./validate";

const target = (address: string, sheetName = "Form") => ({ sheetName, address });

describe("validateMappings (R2.1)", () => {
  it("hợp lệ: một cột gán nhiều ô", () => {
    expect(
      validateMappings(
        [
          { column: 0, target: target("B3") },
          { column: 0, target: target("F30") },
          { column: 1, target: target("A9") },
        ],
        3,
      ),
    ).toEqual([]);
  });

  it("báo lỗi khi chưa có mapping", () => {
    expect(validateMappings([], 3)).toEqual(["Chưa gán cột nào vào ô của form mẫu."]);
  });

  it("báo lỗi khi một ô nhận 2 cột (tên sheet không phân biệt hoa/thường)", () => {
    expect(
      validateMappings(
        [
          { column: 0, target: target("B3") },
          { column: 1, target: target("B3", "FORM") },
        ],
        3,
      ),
    ).toEqual(["Ô FORM!B3 đang được gán cho 2 cột khác nhau."]);
  });

  it("báo lỗi khi cột nằm ngoài vùng nguồn", () => {
    expect(validateMappings([{ column: 5, target: target("B3") }], 3)).toEqual([
      "Cột thứ 6 không nằm trong vùng nguồn.",
    ]);
  });
});
