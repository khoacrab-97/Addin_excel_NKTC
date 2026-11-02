import { describe, expect, it } from "vitest";
import { fillDatePattern, isDatePattern, weekdayName } from "./datePattern";

const D = 46188; // 15/06/2026 (thứ Hai)

describe("fillDatePattern (R2.15)", () => {
  it("rải ngày/tháng/năm vào sau từ khóa", () => {
    expect(fillDatePattern("Ngày ..... tháng ..... năm ..........", D)).toBe("Ngày 15 tháng 06 năm 2026");
    expect(fillDatePattern("Ngày 01 tháng 01 năm 2025", D + 1)).toBe("Ngày 16 tháng 06 năm 2026");
  });

  it("giữ phần chữ phía trước và cách viết hoa", () => {
    expect(fillDatePattern("Hà Nội, ngày … tháng … năm …", D)).toBe("Hà Nội, ngày 15 tháng 06 năm 2026");
    expect(fillDatePattern("NGÀY __ THÁNG __ NĂM ____", D)).toBe("NGÀY 15 THÁNG 06 NĂM 2026");
  });

  it("không phải mẫu ngày → undefined", () => {
    expect(fillDatePattern("Nội dung công việc", D)).toBeUndefined();
    expect(fillDatePattern("năm tháng ngày", D)).toBeUndefined();
    expect(isDatePattern("Ngày ... tháng ... năm ...")).toBe(true);
    expect(isDatePattern("Thời tiết")).toBe(false);
  });
});

describe("weekdayName", () => {
  it("tên thứ trong tuần", () => {
    expect(weekdayName(D)).toBe("Thứ Hai");
    expect(weekdayName(D + 6)).toBe("Chủ nhật");
  });
});
