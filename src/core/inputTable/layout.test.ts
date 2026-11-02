import { describe, expect, it } from "vitest";
import { detectColumns } from "../dailyLog/columns";
import { INPUT_TABLE_HEADERS, buildInputTableLayout, occupiedCells, parseRowCount } from "./layout";

describe("buildInputTableLayout (F4)", () => {
  it("bảng mặc định 10 dòng tại A1", () => {
    const result = buildInputTableLayout("A1", 10);
    expect(result).toMatchObject({
      ok: true,
      layout: {
        address: "A1:D11",
        header: "A1:D1",
        body: "A2:D11",
        columns: { stt: "A2:A11", content: "B2:B11", start: "C2:C11", end: "D2:D11" },
        firstInput: "B2",
        rows: 10,
      },
    });
    expect(result.ok && result.layout.numbers).toEqual([[1], [2], [3], [4], [5], [6], [7], [8], [9], [10]]);
  });

  it("chuẩn hóa ô bắt đầu ($b$3, chữ thường) và qua cột Z", () => {
    expect(buildInputTableLayout("$b$3", 2)).toMatchObject({ ok: true, layout: { address: "B3:E5" } });
    expect(buildInputTableLayout("Y5", 1)).toMatchObject({ ok: true, layout: { address: "Y5:AB6", firstInput: "Z6" } });
  });

  it("báo lỗi ô sai, số dòng sai, vượt giới hạn sheet", () => {
    expect(buildInputTableLayout("B", 10)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("A1:B2", 10)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("A1", 0)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("A1", 1001)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("XFB1", 1)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("A1048576", 1)).toMatchObject({ ok: false });
    expect(buildInputTableLayout("A1048575", 1)).toMatchObject({ ok: true });
  });

  it("tiêu đề được F1 nhận diện đủ 3 cột (Nội dung, Bắt đầu, Kết thúc)", () => {
    expect(detectColumns([...INPUT_TABLE_HEADERS])).toEqual({ content: 1, start: 2, end: 3 });
  });
});

describe("parseRowCount", () => {
  it("nhận số nguyên 1..1000", () => {
    expect(parseRowCount(" 25 ")).toBe(25);
    expect(parseRowCount("1000")).toBe(1000);
    expect(parseRowCount("0")).toBeUndefined();
    expect(parseRowCount("1001")).toBeUndefined();
    expect(parseRowCount("2.5")).toBeUndefined();
    expect(parseRowCount("abc")).toBeUndefined();
  });
});

describe("occupiedCells", () => {
  it("liệt kê ô có dữ liệu hoặc công thức theo địa chỉ trên sheet", () => {
    expect(
      occupiedCells(
        [
          ["", "x"],
          ["", 0],
        ],
        [
          ["=1+1", "x"],
          ["", 0],
        ],
        "C5",
      ),
    ).toEqual(["C5", "D5", "D6"]);
    expect(occupiedCells([["", ""]], [["", ""]], "A1")).toEqual([]);
  });
});
