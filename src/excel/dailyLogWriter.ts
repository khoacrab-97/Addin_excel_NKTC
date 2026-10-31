import type { DailyEntry } from "../core/dailyLog/buildDailyLog";
import { uniqueSheetName } from "../core/dailyLog/sheetName";
import { type DateSystem, toWorkbookSerial } from "../core/dates/excelDate";
import { columnLetter } from "../core/excel/address";

const HEADER = ["Ngày", "Nội dung"] as const;
const DATE_FORMAT = "dd/mm/yyyy";
/** Định dạng Text: nội dung bắt đầu bằng "-", "=", "+" không bị Excel hiểu là công thức. */
const TEXT_FORMAT = "@";
const DATE_COLUMN_WIDTH = 80;
const CONTENT_COLUMN_WIDTH = 480;

/**
 * Tạo sheet mới và ghi nhật ký theo ngày (R1.8, R1.9). Không ghi đè sheet có sẵn.
 * @returns Tên sheet đã tạo.
 */
export async function writeDailyLogSheet(
  entries: readonly DailyEntry[],
  dateSystem: DateSystem = 1900,
): Promise<string> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const sheetName = uniqueSheetName(worksheets.items.map((sheet) => sheet.name));
    const sheet = worksheets.add(sheetName);

    const lastRow = entries.length + 1;
    const header = sheet.getRange("A1:B1");
    header.values = [[...HEADER]];
    header.format.font.bold = true;

    if (entries.length > 0) {
      const dateColumn = sheet.getRange(`A2:A${lastRow}`);
      const contentColumn = sheet.getRange(`B2:B${lastRow}`);
      // Đặt định dạng trước khi ghi giá trị (thứ tự lệnh trong batch được giữ nguyên).
      dateColumn.numberFormat = entries.map(() => [DATE_FORMAT]);
      contentColumn.numberFormat = entries.map(() => [TEXT_FORMAT]);
      dateColumn.values = entries.map((entry) => [toWorkbookSerial(entry.day, dateSystem)]);
      contentColumn.values = entries.map((entry) => [entry.content]);
      contentColumn.format.wrapText = true;
    }

    const table = sheet.getRange(`A1:${columnLetter(HEADER.length - 1)}${lastRow}`);
    table.format.verticalAlignment = "Top";
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
      sheet.getRange("A:A").format.columnWidth = DATE_COLUMN_WIDTH;
      sheet.getRange("B:B").format.columnWidth = CONTENT_COLUMN_WIDTH;
    }

    sheet.activate();
    await context.sync();
    return sheetName;
  });
}
