import { describe, expect, it } from "vitest";
import { columnIndexFromLetter, columnLetter, normalizeCellAddress, stripSheetName } from "./address";

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

describe("columnIndexFromLetter", () => {
  it.each([
    ["A", 0],
    ["z", 25],
    ["AA", 26],
    ["XFD", 16383],
  ])("%s → %i", (letters, expected) => {
    expect(columnIndexFromLetter(letters)).toBe(expected);
  });
});

describe("normalizeCellAddress", () => {
  it.each([
    ["B5", "B5"],
    ["$b$5", "B5"],
    [" c12 ", "C12"],
    ["XFD1048576", "XFD1048576"],
  ])("%j → %s", (address, expected) => {
    expect(normalizeCellAddress(address)).toBe(expected);
  });

  it.each(["B5:C6", "B", "5", "XFE1", "A0", "A1048577", "Form!B5", ""])("từ chối %j", (address) => {
    expect(normalizeCellAddress(address)).toBeUndefined();
  });
});
