import { describe, expect, it } from "vitest";
import { buildDayIndex, detectDateColumn, indexOfDay, stepIndex } from "./days";

const D = 46188; // 15/06/2026

describe("detectDateColumn", () => {
  it("ưu tiên tiêu đề 'Ngày'", () => {
    expect(detectDateColumn([["Nội dung", "NGÀY", "Thời tiết"]])).toBe(1);
  });

  it("nhận tiêu đề chứa từ 'ngày' nhưng không nhầm 'Ngày bắt đầu' khi có cột 'Ngày'", () => {
    expect(detectDateColumn([["Ngày bắt đầu", "Ngày"]])).toBe(1);
    expect(detectDateColumn([["Nội dung", "Ngày thi công"]])).toBe(1);
  });

  it("không có tiêu đề phù hợp → chọn cột đầu tiên toàn ngày hợp lệ", () => {
    expect(
      detectDateColumn([
        ["A", "B"],
        ["x", D],
        ["y", "16/06/2026"],
        ["z", ""],
      ]),
    ).toBe(1);
  });

  it("không nhận diện được → undefined", () => {
    expect(detectDateColumn([["A"], ["x"]])).toBeUndefined();
  });
});

describe("buildDayIndex (R2.10)", () => {
  it("sắp xếp tăng dần, giữ số dòng sheet, bỏ qua ô ngày trống, báo ngày sai và ngày trùng", () => {
    const values = [
      ["Ngày", "Nội dung"],
      [D + 1, "b"],
      [D, "a"],
      ["", "ghi chú"],
      ["abc", "x"],
      [D, "trùng"],
    ];
    const result = buildDayIndex(values, 0, 3);
    expect(result.days).toEqual([
      { day: D, valueRowIndex: 2, rowNumber: 5 },
      { day: D + 1, valueRowIndex: 1, rowNumber: 4 },
    ]);
    expect(result.issues).toEqual([
      { rowNumber: 7, reason: "Ngày không hợp lệ." },
      { rowNumber: 8, reason: "Trùng ngày 15/06/2026 với dòng 5; bỏ qua dòng này." },
    ]);
  });
});

describe("điều hướng", () => {
  const days = [
    { day: D, valueRowIndex: 1, rowNumber: 2 },
    { day: D + 2, valueRowIndex: 2, rowNumber: 3 },
  ];

  it("indexOfDay", () => {
    expect(indexOfDay(days, D + 2)).toBe(1);
    expect(indexOfDay(days, D + 1)).toBe(-1);
  });

  it("stepIndex giới hạn trong phạm vi", () => {
    expect(stepIndex(2, 0, 1)).toBe(1);
    expect(stepIndex(2, 1, 1)).toBe(1);
    expect(stepIndex(2, 0, -1)).toBe(0);
    expect(stepIndex(2, 0, Infinity)).toBe(1);
    expect(stepIndex(2, 1, -Infinity)).toBe(0);
    expect(stepIndex(0, 0, 1)).toBe(-1);
  });
});
