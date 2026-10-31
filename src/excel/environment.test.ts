import { afterEach, describe, expect, it, vi } from "vitest";
import { detectMaxExcelApi, getHostEnvironment } from "./environment";

/** Giả lập `Office` toàn cục với mức ExcelApi tối đa cho trước (0 = không hỗ trợ). */
function stubOffice(maxMinor: number, diagnostics?: { platform: string; version: string }): void {
  vi.stubGlobal("Office", {
    context: {
      diagnostics,
      requirements: {
        isSetSupported: (name: string, version: string) =>
          name === "ExcelApi" && Number(version.split(".")[1]) <= maxMinor,
      },
    },
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("detectMaxExcelApi", () => {
  it("trả về 1.1 với Excel 2016 (bản mua một lần)", () => {
    stubOffice(1);
    expect(detectMaxExcelApi()).toBe("1.1");
  });

  it("trả về mức cao nhất được hỗ trợ với Microsoft 365", () => {
    stubOffice(17);
    expect(detectMaxExcelApi()).toBe("1.17");
  });

  it("báo không hỗ trợ khi không có ExcelApi", () => {
    stubOffice(0);
    expect(detectMaxExcelApi()).toBe("Không hỗ trợ");
  });
});

describe("getHostEnvironment", () => {
  it("đọc nền tảng và phiên bản Office từ diagnostics", () => {
    stubOffice(1, { platform: "PC", version: "16.0.4266.1001" });
    expect(getHostEnvironment()).toEqual({ platform: "PC", officeVersion: "16.0.4266.1001", maxExcelApi: "1.1" });
  });

  it("không lỗi khi bản Office cũ không có diagnostics", () => {
    stubOffice(1);
    expect(getHostEnvironment()).toEqual({
      platform: "Không xác định",
      officeVersion: "Không xác định",
      maxExcelApi: "1.1",
    });
  });
});
