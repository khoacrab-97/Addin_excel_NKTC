import { describe, expect, it } from "vitest";
import { normalizeHeader, removeVietnameseDiacritics } from "./normalize";

describe("removeVietnameseDiacritics", () => {
  it("bỏ dấu cả chữ hoa và chữ thường", () => {
    expect(removeVietnameseDiacritics("Ngày Bắt Đầu")).toBe("ngay bat dau");
    expect(removeVietnameseDiacritics("NỘI DUNG CÔNG VIỆC")).toBe("noi dung cong viec");
    expect(removeVietnameseDiacritics("Thời tiết, Ước lượng, Kỹ thuật")).toBe("thoi tiet, uoc luong, ky thuat");
  });
});

describe("normalizeHeader", () => {
  it("gộp khoảng trắng, bỏ ký tự đặc biệt và xuống dòng", () => {
    expect(normalizeHeader("  Ngày\nbắt   đầu (*) ")).toBe("ngay bat dau");
    expect(normalizeHeader("Ban chỉ huy - CBKT")).toBe("ban chi huy cbkt");
  });
});
