import { describe, expect, it } from "vitest";
import { isDialogView, parsePayload, parseReply, shortLabel } from "./protocol";

describe("parsePayload", () => {
  it("nhận payload hợp lệ", () => {
    const payload = {
      source: { value: "NK!A1:C9" },
      options: [{ index: 0, label: "Cột A – Nội dung" }],
      detected: { content: 0 },
    };
    expect(parsePayload("f1-columns", JSON.stringify(payload))).toEqual(payload);
    expect(parsePayload("help", "{}")).toEqual({});
  });

  it("từ chối JSON hỏng hoặc sai kiểu", () => {
    expect(parsePayload("message", "{")).toBeUndefined();
    expect(parsePayload("message", JSON.stringify({ title: "x", tone: "loud", lines: [] }))).toBeUndefined();
    expect(
      parsePayload(
        "export-form",
        JSON.stringify({
          source: { value: "x" },
          options: [],
          dayCount: 0,
          state: { rows: [{ column: -1, target: "" }] },
        }),
      ),
    ).toBeUndefined();
    expect(
      parsePayload("message", JSON.stringify({ title: "x", tone: "info", lines: [], link: "https://evil.test" })),
    ).toBeUndefined();
    expect(parsePayload("confirm", JSON.stringify([]))).toBeUndefined();
  });
});

describe("parseReply", () => {
  it("Tạo bảng: nhận lựa chọn hợp lệ, từ chối số dòng/vị trí sai", () => {
    expect(
      parseReply("create-table", JSON.stringify({ action: "create", target: "new", cell: "A1", rows: 10 })),
    ).toEqual({ action: "create", target: "new", cell: "A1", rows: 10 });
    expect(parseReply("create-table", JSON.stringify({ action: "create", target: "x", cell: "A1", rows: 10 }))).toEqual(
      {
        action: "cancel",
      },
    );
    expect(
      parseReply("create-table", JSON.stringify({ action: "create", target: "current", cell: "A1", rows: 0 })),
    ).toEqual({ action: "cancel" });
  });

  it("nhận kết quả hợp lệ", () => {
    const state = { dateColumn: 0, rows: [{ column: 1, target: "Mẫu NKTC!E2" }, { target: "" }] };
    expect(parseReply("export-form", JSON.stringify({ action: "assign", state }))).toEqual({ action: "assign", state });
    expect(parseReply("export-form", JSON.stringify({ action: "pick-target", row: 1, state }))).toEqual({
      action: "pick-target",
      row: 1,
      state,
    });
    expect(parseReply("f1-columns", JSON.stringify({ action: "pick", address: "TienDo!A1:D11" }))).toEqual({
      action: "pick",
      address: "TienDo!A1:D11",
    });
  });

  it("dữ liệu không hợp lệ được coi là hủy", () => {
    expect(
      parseReply("export-form", JSON.stringify({ action: "assign", state: { rows: [{ column: 1.5, target: "" }] } })),
    ).toEqual({
      action: "cancel",
    });
    expect(parseReply("f1-columns", JSON.stringify({ action: "ok", columns: { content: 0 } }))).toEqual({
      action: "cancel",
    });
    expect(parseReply("goto", "not json")).toEqual({ action: "cancel" });
  });
});

describe("tiện ích", () => {
  it("isDialogView", () => {
    expect(isDialogView("goto")).toBe(true);
    expect(isDialogView("taskpane")).toBe(false);
  });

  it("shortLabel gọn khoảng trắng và cắt nhãn dài", () => {
    expect(shortLabel("  Nội   dung \n công việc ")).toBe("Nội dung công việc");
    expect(shortLabel("x".repeat(100))).toHaveLength(60);
  });
});
