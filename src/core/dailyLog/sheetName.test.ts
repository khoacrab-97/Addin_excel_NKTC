import { describe, expect, it } from "vitest";
import { DAILY_SHEET_BASE_NAME, uniqueSheetName } from "./sheetName";

describe("uniqueSheetName (R1.9)", () => {
  it("dùng tên mặc định khi chưa tồn tại", () => {
    expect(uniqueSheetName(["Sheet1"])).toBe(DAILY_SHEET_BASE_NAME);
  });

  it("thêm hậu tố khi trùng, không phân biệt hoa/thường", () => {
    expect(uniqueSheetName(["nk_theongay"])).toBe("NK_TheoNgay_2");
    expect(uniqueSheetName(["NK_TheoNgay", "NK_TheoNgay_2", "NK_TheoNgay_4"])).toBe("NK_TheoNgay_3");
  });
});
