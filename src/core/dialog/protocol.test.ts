import { describe, expect, it } from "vitest";
import { isDialogView, parsePayload, parseReply, shortLabel } from "./protocol";

describe("parsePayload", () => {
  it("nhận payload hợp lệ", () => {
    const payload = { source: "NK!A1:C9", options: [{ index: 0, label: "A – Nội dung" }], detected: { content: 0 } };
    expect(parsePayload("f1-columns", JSON.stringify(payload))).toEqual(payload);
    expect(parsePayload("help", "{}")).toEqual({});
  });

  it("từ chối JSON hỏng hoặc sai kiểu", () => {
    expect(parsePayload("message", "{")).toBeUndefined();
    expect(parsePayload("message", JSON.stringify({ title: "x", tone: "loud", lines: [] }))).toBeUndefined();
    expect(parsePayload("assign", JSON.stringify({ target: "A1", source: "x", options: [], dateColumn: -1 }))).toBe(
      undefined,
    );
    expect(parsePayload("confirm", JSON.stringify([]))).toBeUndefined();
  });
});

describe("parseReply", () => {
  it("nhận kết quả hợp lệ", () => {
    expect(parseReply("assign", JSON.stringify({ action: "assign", column: 3 }))).toEqual({
      action: "assign",
      column: 3,
    });
    expect(parseReply("assign", JSON.stringify({ action: "unassign" }))).toEqual({ action: "unassign" });
    expect(parseReply("settings", JSON.stringify({ action: "save", dateColumn: 0, remove: ["a!B5"] }))).toEqual({
      action: "save",
      dateColumn: 0,
      remove: ["a!B5"],
    });
  });

  it("dữ liệu không hợp lệ được coi là hủy", () => {
    expect(parseReply("assign", JSON.stringify({ action: "assign", column: 1.5 }))).toEqual({ action: "cancel" });
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
