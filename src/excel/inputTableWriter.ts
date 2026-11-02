import { uniqueSheetName } from "../core/dailyLog/sheetName";
import {
  INPUT_SHEET_BASE_NAME,
  INPUT_TABLE_HEADERS,
  type InputTableLayout,
  occupiedCells,
} from "../core/inputTable/layout";
import { UserFacingError } from "./errors";

const DATE_FORMAT = "dd/mm/yyyy";
const TEXT_FORMAT = "@";
const TABLE_STYLE = "TableStyleMedium2";
/** Độ rộng cột (point): STT, Nội dung, Bắt đầu, Kết thúc — chỉ đặt khi có ExcelApi 1.2. */
const COLUMN_WIDTHS = [36, 360, 80, 80] as const;
const MAX_LISTED_CELLS = 8;

export type TableTarget = { kind: "current" } | { kind: "new" };

export interface CreatedTable {
  sheetName: string;
  address: string;
}

/** Sheet và ô đang chọn (ô trên-trái nếu chọn nhiều ô) — gợi ý vị trí bảng. */
export async function getActiveCell(): Promise<{ sheetName: string; cell: string }> {
  return Excel.run(async (context) => {
    const topLeft = context.workbook.getSelectedRange().getCell(0, 0);
    topLeft.load("address");
    topLeft.worksheet.load("name");
    await context.sync();
    const address = topLeft.address;
    return { sheetName: topLeft.worksheet.name, cell: address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, "") };
  });
}

/**
 * Tạo bảng nhập tiến độ (F4): Excel Table 4 cột có tiêu đề, STT 1..n, định dạng ngày cho Bắt đầu/Kết thúc.
 * Sheet hiện tại: chỉ tạo khi toàn bộ vùng bảng còn trống (R4.4) — không ghi đè dữ liệu/công thức.
 * Sheet mới: tên `TienDo`, `TienDo_2`… (không trùng sheet có sẵn).
 */
export async function createInputTable(target: TableTarget, layout: InputTableLayout): Promise<CreatedTable> {
  return Excel.run(async (context) => {
    let sheet: Excel.Worksheet;
    if (target.kind === "new") {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      sheet = worksheets.add(
        uniqueSheetName(
          worksheets.items.map((item) => item.name),
          INPUT_SHEET_BASE_NAME,
        ),
      );
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
      const area = sheet.getRange(layout.address);
      area.load(["values", "formulas"]);
      await context.sync();
      const occupied = occupiedCells(
        area.values as unknown[][],
        area.formulas as unknown[][],
        layout.address.split(":")[0] ?? "A1",
      );
      if (occupied.length > 0) {
        const listed = occupied.slice(0, MAX_LISTED_CELLS).join(", ");
        const more = occupied.length > MAX_LISTED_CELLS ? ` và ${occupied.length - MAX_LISTED_CELLS} ô khác` : "";
        throw new UserFacingError(
          `Vùng ${layout.address} đang có dữ liệu (${listed}${more}). Add-in không ghi đè — chọn ô bắt đầu khác hoặc tạo trên sheet mới.`,
        );
      }
    }
    sheet.load("name");

    // Định dạng trước khi ghi giá trị (thứ tự lệnh trong batch được giữ nguyên).
    const columns = layout.columns;
    sheet.getRange(columns.content).numberFormat = layout.numbers.map(() => [TEXT_FORMAT]);
    sheet.getRange(columns.start).numberFormat = layout.numbers.map(() => [DATE_FORMAT]);
    sheet.getRange(columns.end).numberFormat = layout.numbers.map(() => [DATE_FORMAT]);
    sheet.getRange(columns.content).format.wrapText = true;
    sheet.getRange(layout.header).values = [[...INPUT_TABLE_HEADERS]];
    sheet.getRange(columns.stt).values = layout.numbers;

    const table = sheet.tables.add(layout.address, true);
    table.style = TABLE_STYLE;
    sheet.getRange(layout.address).format.verticalAlignment = "Top";
    sheet.getRange(columns.stt).format.horizontalAlignment = "Center";

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
      [columns.stt, columns.content, columns.start, columns.end].forEach((address, index) => {
        sheet.getRange(address).format.columnWidth = COLUMN_WIDTHS[index] ?? 80;
      });
    }
    sheet.activate();
    sheet.getRange(layout.firstInput).select();
    try {
      await context.sync();
    } catch (error) {
      const code = typeof error === "object" && error !== null && "code" in error ? String(error.code) : "";
      if (code === "InvalidOperation" || code === "InvalidArgument") {
        throw new UserFacingError(
          `Không tạo được bảng tại ${layout.address} (vùng có thể chồng lên bảng/ô merge có sẵn hoặc sheet đang được bảo vệ).`,
        );
      }
      throw error;
    }
    return { sheetName: sheet.name, address: layout.address };
  });
}
