import { normalizeCellAddress, stripSheetName } from "../core/address/address";
import type { TargetCell } from "../core/mapping/types";
import { UserFacingError } from "./errors";
import { assertSheetsExist } from "./worksheets";

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
 * Vùng dữ liệu cho lệnh trên Ribbon: vùng đang chọn; nếu chỉ chọn 1 ô thì lấy toàn bộ vùng có dữ liệu
 * của sheet (used range) — người dùng chỉ cần đặt con trỏ trong bảng.
 */
export async function getSelectedTableRef(): Promise<RangeRef> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "columnCount"]);
    range.worksheet.load("name");
    await context.sync();
    if (range.address.includes(",")) {
      throw new UserFacingError("Vui lòng chỉ chọn một vùng dữ liệu liên tục.");
    }
    if (range.rowCount > 1 || range.columnCount > 1) {
      return { sheetName: range.worksheet.name, address: stripSheetName(range.address) };
    }
    const used = range.worksheet.getUsedRange();
    used.load("address");
    await context.sync();
    return { sheetName: range.worksheet.name, address: stripSheetName(used.address) };
  });
}

/**
 * Đọc giá trị mới nhất của vùng (CLAUDE.md §4: không dùng dữ liệu cache).
 * Vùng được cắt theo used range để tránh tải hàng triệu ô khi người dùng chọn nguyên cột.
 */
export async function readRange(ref: RangeRef): Promise<RangeSnapshot> {
  return Excel.run(async (context) => {
    await assertSheetsExist(context, [ref.sheetName]);
    const sheet = context.workbook.worksheets.getItem(ref.sheetName);
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

/**
 * Lấy ô đang chọn làm ô đích (R2.2): nếu chọn ô merge (hoặc nhiều ô), lấy ô đầu tiên (trên-trái).
 * Dùng `getCell(0, 0)` vì `getMergedAreas` cần ExcelApi 1.13.
 */
export async function getSelectedCell(): Promise<TargetCell> {
  return Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange();
    const topLeft = selection.getCell(0, 0);
    selection.load("address");
    topLeft.load("address");
    topLeft.worksheet.load("name");
    await context.sync();
    if (selection.address.includes(",")) {
      throw new UserFacingError("Vui lòng chỉ chọn một ô (hoặc một ô merge) trên form mẫu.");
    }
    const address = normalizeCellAddress(stripSheetName(topLeft.address));
    if (address === undefined) throw new UserFacingError("Không xác định được địa chỉ ô đang chọn.");
    return { sheetName: topLeft.worksheet.name, address };
  });
}
