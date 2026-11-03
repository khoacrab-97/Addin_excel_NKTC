import { describe, expect, it } from "vitest";
import { remapFormState } from "./formRemap";

const HEADERS = ["Ngày", "Nội dung", "Thời tiết", "Ghi chú", "Ghi chú"];

describe("remapFormState", () => {
  const state = {
    dateColumn: 0,
    selected: 1,
    rows: [{ column: 0, target: "Mẫu NKTC!E2" }, { column: 2, target: "Mẫu NKTC!C6" }, { target: "" }],
  };

  it("chèn cột → dòng gán theo tên; dòng chưa chọn cột giữ nguyên", () => {
    expect(remapFormState(state, HEADERS, ["STT", ...HEADERS])).toEqual({
      state: {
        dateColumn: 1,
        selected: 1,
        rows: [{ column: 1, target: "Mẫu NKTC!E2" }, { column: 3, target: "Mẫu NKTC!C6" }, { target: "" }],
      },
      lost: [],
    });
  });

  it("cột không còn → bỏ chọn cột, báo theo dòng; cột Ngày không còn → báo", () => {
    const result = remapFormState(state, HEADERS, ["Ngày tháng", "Nội dung"]);
    expect(result.state).toEqual({
      selected: 1,
      rows: [{ target: "Mẫu NKTC!E2" }, { target: "Mẫu NKTC!C6" }, { target: "" }],
    });
    expect(result.lost).toEqual([
      'Dòng 1: cột "Ngày" không còn trong vùng nguồn (bị xóa hoặc đổi tên) — chọn lại cột.',
      'Dòng 2: cột "Thời tiết" không còn trong vùng nguồn (bị xóa hoặc đổi tên) — chọn lại cột.',
      'Cột Ngày "Ngày" không còn trong vùng nguồn — chọn lại cột ngày.',
    ]);
  });
});
