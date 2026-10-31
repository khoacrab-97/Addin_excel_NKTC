import { stripSheetName } from "../core/excel/address";
import { UserFacingError } from "./errors";

/** Tham chiếu tới một vùng dữ liệu để đọc lại về sau. */
export interface RangeRef {
  sheetName: string;
  /** Địa chỉ không gồm tên sheet, ví dụ `A1:D50`. */
  address: string;
}

/** Giá trị của vùng (đã giới hạn trong used range) và vị trí của nó trên sheet. */
export interface RangeSnapshot {
  ref: RangeRef;
  values: unknown[][];
  /** Số dòng trên sheet (1-based) của dòng đầu tiên. */
  firstRowNumber: number;
  /** Chỉ số cột (0-based) của cột đầu tiên. */
  firstColumnIndex: number;
}

/** Lấy tham chiếu vùng đang được chọn trên sheet. */
export async function getSelectedRangeRef(): Promise<RangeRef> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    range.worksheet.load("name");
    await context.sync();
    // Chọn nhiều vùng rời nhau (giữ Ctrl) cho địa chỉ dạng "A1:B2,D1:D5".
    if (range.address.includes(",")) {
      throw new UserFacingError("Vui lòng chỉ chọn một vùng dữ liệu liên tục.");
    }
    return { sheetName: range.worksheet.name, address: stripSheetName(range.address) };
  });
}

/**
 * Đọc giá trị mới nhất của vùng (CLAUDE.md §4: không dùng dữ liệu cache).
 * Vùng được cắt theo used range để tránh tải hàng triệu ô khi người dùng chọn nguyên cột.
 */
export async function readRange(ref: RangeRef): Promise<RangeSnapshot> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    // ExcelApi 1.1 không có getItemOrNullObject → kiểm tra tên sheet thủ công.
    if (!worksheets.items.some((sheet) => sheet.name === ref.sheetName)) {
      throw new UserFacingError(`Không tìm thấy sheet "${ref.sheetName}". Sheet có thể đã bị xóa hoặc đổi tên.`);
    }

    const sheet = worksheets.getItem(ref.sheetName);
    const target = sheet.getRange(ref.address).getIntersection(sheet.getUsedRange());
    target.load(["address", "values", "rowIndex", "columnIndex"]);
    try {
      await context.sync();
    } catch (error) {
      if (typeof error === "object" && error !== null && "code" in error && error.code === "ItemNotFound") {
        throw new UserFacingError(`Vùng ${ref.address} trên sheet "${ref.sheetName}" không có dữ liệu.`);
      }
      throw error;
    }

    return {
      ref: { sheetName: ref.sheetName, address: stripSheetName(target.address) },
      values: target.values as unknown[][],
      firstRowNumber: target.rowIndex + 1,
      firstColumnIndex: target.columnIndex,
    };
  });
}
