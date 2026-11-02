import { describe, expect, it } from "vitest";
import type { SourceDay } from "./days";
import { resolveDay } from "./navigation";

const D = 46188; // 15/06/2026
const days: SourceDay[] = [
  { day: D, valueRowIndex: 1, rowNumber: 2 },
  { day: D + 1, valueRowIndex: 2, rowNumber: 3 },
  { day: D + 3, valueRowIndex: 3, rowNumber: 4 },
];

describe("resolveDay", () => {
  it("Đầu / Cuối", () => {
    expect(resolveDay(days, undefined, { kind: "first" })).toEqual({ ok: true, target: days[0] });
    expect(resolveDay(days, D, { kind: "last" })).toEqual({ ok: true, target: days[2] });
  });

  it("Trước / Sau theo ngày đang hiển thị (bỏ qua ngày không có dữ liệu)", () => {
    expect(resolveDay(days, D + 1, { kind: "next" })).toEqual({ ok: true, target: days[2] });
    expect(resolveDay(days, D + 3, { kind: "prev" })).toEqual({ ok: true, target: days[1] });
  });

  it("chưa có ngày đang hiển thị → Trước/Sau hiển thị ngày đầu, không thông báo", () => {
    expect(resolveDay(days, undefined, { kind: "next" })).toEqual({ ok: true, target: days[0] });
    expect(resolveDay(days, undefined, { kind: "prev" })).toEqual({ ok: true, target: days[0] });
  });

  it("ngày đang hiển thị không còn trong nguồn → về ngày đầu kèm thông báo", () => {
    const result = resolveDay(days, D + 2, { kind: "next" });
    expect(result).toMatchObject({ ok: true, target: days[0] });
    expect(result.ok && result.notice).toContain("17/06/2026");
  });

  it("đã ở đầu/cuối → không ghi, báo thông tin", () => {
    expect(resolveDay(days, D + 3, { kind: "next" })).toEqual({
      ok: false,
      severity: "info",
      message: "Đang ở ngày cuối cùng (18/06/2026).",
    });
    expect(resolveDay(days, D, { kind: "prev" })).toMatchObject({ ok: false, severity: "info" });
  });

  it("Đến ngày: có / không có trong nguồn", () => {
    expect(resolveDay(days, undefined, { kind: "goto", day: D + 1 })).toEqual({ ok: true, target: days[1] });
    expect(resolveDay(days, undefined, { kind: "goto", day: D + 2 })).toEqual({
      ok: false,
      severity: "error",
      message: "Ngày 17/06/2026 không có trong vùng nguồn (từ 15/06/2026 đến 18/06/2026).",
    });
  });

  it("nguồn không có ngày → lỗi", () => {
    expect(resolveDay([], undefined, { kind: "first" })).toMatchObject({ ok: false, severity: "error" });
  });
});
