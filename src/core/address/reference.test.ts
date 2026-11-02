import { describe, expect, it } from "vitest";
import { formatReference, parseCellReference, parseRangeReference } from "./reference";

describe("parseRangeReference", () => {
  it("đọc vùng có/không có tên sheet, tên sheet có dấu, khoảng trắng, dấu nháy", () => {
    expect(parseRangeReference("NK_TheoNgay!A1:I201")).toEqual({ sheetName: "NK_TheoNgay", address: "A1:I201" });
    expect(parseRangeReference("Mẫu NKTC!$e$2")).toEqual({ sheetName: "Mẫu NKTC", address: "E2" });
    expect(parseRangeReference("'Tiến độ ''A'''!a:c")).toEqual({ sheetName: "Tiến độ 'A'", address: "A:C" });
    expect(parseRangeReference(" b2:d10 ")).toEqual({ address: "B2:D10" });
  });

  it("từ chối địa chỉ sai, tên sheet sai", () => {
    expect(parseRangeReference("")).toBeUndefined();
    expect(parseRangeReference("A1,B2")).toBeUndefined();
    expect(parseRangeReference("Sheet[1]!A1")).toBeUndefined();
    expect(parseRangeReference("!A1")).toBeUndefined();
    expect(parseRangeReference(`${"x".repeat(32)}!A1`)).toBeUndefined();
    expect(parseRangeReference("Sheet1!ZZZZ1")).toBeUndefined();
  });
});

describe("parseCellReference", () => {
  it("chỉ nhận một ô", () => {
    expect(parseCellReference("Mẫu NKTC!e2")).toEqual({ sheetName: "Mẫu NKTC", address: "E2" });
    expect(parseCellReference("C6")).toEqual({ address: "C6" });
    expect(parseCellReference("Mẫu NKTC!E2:I2")).toBeUndefined();
  });

  it("formatReference", () => {
    expect(formatReference("Mẫu NKTC", "E2")).toBe("Mẫu NKTC!E2");
  });
});
