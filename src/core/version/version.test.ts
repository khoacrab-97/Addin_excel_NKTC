import { describe, expect, it } from "vitest";
import { compareVersions, evaluateUpdate, parseReleaseInfo, parseVersion } from "./version";

describe("parseVersion / compareVersions", () => {
  it("đọc x.y.z", () => {
    expect(parseVersion("1.2.3")).toEqual([1, 2, 3]);
    expect(parseVersion(" 10.0.12 ")).toEqual([10, 0, 12]);
    expect(parseVersion("1.2")).toBeUndefined();
    expect(parseVersion("1.2.3-beta")).toBeUndefined();
  });

  it("so sánh theo số, không theo chuỗi", () => {
    expect(compareVersions("1.0.0", "1.0.0")).toBe(0);
    expect(compareVersions("1.0.0", "1.0.1")).toBeLessThan(0);
    expect(compareVersions("1.10.0", "1.9.9")).toBeGreaterThan(0);
    expect(compareVersions("2.0.0", "10.0.0")).toBeLessThan(0);
    expect(compareVersions("abc", "1.0.0")).toBeUndefined();
  });
});

describe("parseReleaseInfo", () => {
  it("đọc version.json hợp lệ", () => {
    expect(parseReleaseInfo('{"version":"1.2.0","builtAt":"2026-10-31T06:00:00Z"}')).toEqual({
      version: "1.2.0",
      builtAt: "2026-10-31T06:00:00Z",
    });
    expect(parseReleaseInfo('{"version":"1.2.0"}')).toEqual({ version: "1.2.0" });
  });

  it.each(["", "not json", "null", "[]", '{"version":1}', '{"version":"1.2"}', "<html>404</html>"])(
    "từ chối %j",
    (text) => {
      expect(parseReleaseInfo(text)).toBeUndefined();
    },
  );
});

describe("evaluateUpdate (P3.2)", () => {
  it("phân loại trạng thái cập nhật", () => {
    expect(evaluateUpdate("1.0.0", { version: "1.0.0" })).toEqual({
      kind: "latest",
      current: "1.0.0",
      latest: "1.0.0",
    });
    expect(evaluateUpdate("1.0.0", { version: "1.1.0" })).toEqual({
      kind: "available",
      current: "1.0.0",
      latest: "1.1.0",
    });
    expect(evaluateUpdate("1.2.0", { version: "1.1.0" })).toMatchObject({ kind: "newer-than-release" });
    expect(evaluateUpdate("1.0.0", undefined)).toEqual({ kind: "unknown", current: "1.0.0" });
    expect(evaluateUpdate("dev", { version: "1.0.0" })).toEqual({ kind: "unknown", current: "dev" });
  });
});
