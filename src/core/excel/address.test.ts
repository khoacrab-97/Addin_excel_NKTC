import { describe, expect, it } from "vitest";
import { columnLetter, stripSheetName } from "./address";

describe("stripSheetName", () => {
  it.each([
    ["Sheet1!A1:C10", "A1:C10"],
    ["'Nhật ký thi công'!B2:D50", "B2:D50"],
    ["'It''s!ok'!A1", "A1"],
    ["A1:C10", "A1:C10"],
  ])("%s → %s", (address, expected) => {
    expect(stripSheetName(address)).toBe(expected);
  });
});

describe("columnLetter", () => {
  it.each([
    [0, "A"],
    [25, "Z"],
    [26, "AA"],
    [701, "ZZ"],
    [702, "AAA"],
    [16383, "XFD"],
  ])("%i → %s", (index, expected) => {
    expect(columnLetter(index)).toBe(expected);
  });

  it("từ chối chỉ số âm", () => {
    expect(() => columnLetter(-1)).toThrow(RangeError);
  });
});
