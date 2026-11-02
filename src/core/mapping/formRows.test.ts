import { describe, expect, it } from "vitest";
import { blankHeaderColumns, buildMappingsFromRows } from "./formRows";

describe("buildMappingsFromRows (Xuất vào form)", () => {
  it("chuyển dòng gán thành mapping; ô không ghi sheet thuộc sheet form mặc định; bỏ dòng trống", () => {
    expect(
      buildMappingsFromRows(
        [{ column: 0, target: "Mẫu NKTC!$E$2" }, { target: "" }, { column: 1, target: "a21" }],
        3,
        "Mẫu NKTC",
      ),
    ).toEqual({
      ok: true,
      mappings: [
        { column: 0, target: { sheetName: "Mẫu NKTC", address: "E2" } },
        { column: 1, target: { sheetName: "Mẫu NKTC", address: "A21" } },
      ],
    });
  });

  it("báo lỗi theo dòng: thiếu cột, thiếu ô, ô sai, ô trùng (một ô một cột)", () => {
    const result = buildMappingsFromRows(
      [
        { target: "E2" },
        { column: 1, target: "" },
        { column: 1, target: "E2:I2" },
        { column: 0, target: "Form!B5" },
        { column: 2, target: "form!b5" },
        { column: 9, target: "C1" },
      ],
      3,
      "Form",
    );
    expect(result).toEqual({
      ok: false,
      errors: [
        "Dòng 1: chưa chọn cột nguồn.",
        "Dòng 2: chưa chọn ô trên form.",
        'Dòng 3: "E2:I2" không phải một ô hợp lệ (ví dụ Mẫu NKTC!E2).',
        "Dòng 5: ô form!B5 đã được gán ở dòng 4.",
        "Dòng 6: chưa chọn cột nguồn.",
      ],
    });
  });

  it("không có dòng nào → yêu cầu thêm", () => {
    expect(buildMappingsFromRows([{ target: " " }], 3, "F")).toMatchObject({ ok: false });
  });
});

describe("blankHeaderColumns", () => {
  it("liệt kê cột tiêu đề trống", () => {
    expect(blankHeaderColumns(["Ngày", "", "Thời tiết", " "])).toEqual([1, 3]);
  });
});
