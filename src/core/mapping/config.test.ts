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
  ])("từ chối khi %s", (_label, raw) => {
    expect(parseMappingConfig(raw)).toBeUndefined();
  });
});

describe("headerTexts", () => {
  it("chuyển tiêu đề thành chuỗi", () => {
    expect(headerTexts(["Ngày", 1, true, null, ""])).toEqual(["Ngày", "1", "true", "", ""]);
  });
});
