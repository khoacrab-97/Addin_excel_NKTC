import { describe, expect, it } from "vitest";
import { headerTexts, parseMappingConfig } from "./config";

const valid = {
  version: 1,
  source: { sheetName: "NK_TheoNgay", address: "A1:G201" },
  headers: ["Ngày", "Nội dung"],
  dateColumn: 0,
  mappings: [{ column: 1, target: { sheetName: "Form", address: "$a$9" } }],
  currentDay: 46188,
};

describe("parseMappingConfig (R2.8)", () => {
  it("đọc cấu hình hợp lệ và chuẩn hóa địa chỉ ô", () => {
    expect(parseMappingConfig(valid)).toEqual({
      ...valid,
      mappings: [{ column: 1, target: { sheetName: "Form", address: "A9" } }],
    });
  });

  it("bỏ currentDay không hợp lệ nhưng vẫn giữ cấu hình", () => {
    expect(parseMappingConfig({ ...valid, currentDay: "x" })).not.toHaveProperty("currentDay");
  });

  it("chấp nhận cấu hình đúng giới hạn (500 mapping, tên sheet 31 ký tự, vùng nguyên cột)", () => {
    const atLimit = {
      ...valid,
      source: { sheetName: "S".repeat(31), address: "A:I" },
      mappings: Array.from({ length: 500 }, (_, i) => ({
        column: 1,
        target: { sheetName: "F".repeat(31), address: `A${String(i + 1)}` },
      })),
    };
    expect(parseMappingConfig(atLimit)?.mappings).toHaveLength(500);
  });

  it("bỏ currentDay vượt ngày lớn nhất Excel", () => {
    expect(parseMappingConfig({ ...valid, currentDay: 3_000_000 })).not.toHaveProperty("currentDay");
  });

  it.each([
    ["không phải object", "abc"],
    ["sai version", { ...valid, version: 2 }],
    ["thiếu source", { ...valid, source: undefined }],
    ["headers không phải chuỗi", { ...valid, headers: [1] }],
    ["dateColumn âm", { ...valid, dateColumn: -1 }],
    [
      "ô đích là vùng nhiều ô",
      { ...valid, mappings: [{ column: 1, target: { sheetName: "Form", address: "A9:B10" } }] },
    ],
    ["mapping thiếu sheet", { ...valid, mappings: [{ column: 1, target: { address: "A9" } }] }],
    ["địa chỉ vùng nguồn không hợp lệ", { ...valid, source: { sheetName: "NK_TheoNgay", address: "=HYPERLINK(1)" } }],
    ["địa chỉ vùng nguồn gồm tên sheet", { ...valid, source: { sheetName: "NK_TheoNgay", address: "Khac!A1:B2" } }],
    ["tên sheet nguồn dài hơn 31 ký tự", { ...valid, source: { sheetName: "S".repeat(32), address: "A1:B2" } }],
    [
      "tên sheet đích dài hơn 31 ký tự",
      { ...valid, mappings: [{ column: 1, target: { sheetName: "F".repeat(32), address: "A9" } }] },
    ],
    [
      "quá nhiều mapping",
      {
        ...valid,
        mappings: Array.from({ length: 501 }, (_, i) => ({
          column: 1,
          target: { sheetName: "Form", address: `A${String(i + 1)}` },
        })),
      },
    ],
    ["quá nhiều tiêu đề", { ...valid, headers: Array.from({ length: 16_385 }, () => "x") }],
    ["tiêu đề dài hơn giới hạn ô", { ...valid, headers: ["x".repeat(32_768)] }],
    [
      "cột vượt số cột Excel",
      { ...valid, mappings: [{ column: 16_384, target: { sheetName: "Form", address: "A9" } }] },
    ],
    ["cột Ngày vượt số cột Excel", { ...valid, dateColumn: 16_384 }],
  ])("từ chối khi %s", (_label, raw) => {
    expect(parseMappingConfig(raw)).toBeUndefined();
  });
});

describe("confirmedKey (R2.5)", () => {
  it("giữ xác nhận khi khớp bộ ô đích hiện tại", () => {
    expect(parseMappingConfig({ ...valid, confirmedKey: "form!A9" })).toHaveProperty("confirmedKey", "form!A9");
  });

  it("bỏ xác nhận không khớp bộ ô đích (không mở rộng quyền ghi)", () => {
    expect(parseMappingConfig({ ...valid, confirmedKey: "form!A9|form!B1" })).not.toHaveProperty("confirmedKey");
    expect(parseMappingConfig({ ...valid, confirmedKey: 1 })).not.toHaveProperty("confirmedKey");
  });
});

describe("headerTexts", () => {
  it("chuyển tiêu đề thành chuỗi", () => {
    expect(headerTexts(["Ngày", 1, true, null, ""])).toEqual(["Ngày", "1", "true", "", ""]);
  });
});
