import { UserFacingError } from "./errors";

/**
 * Kiểm tra các sheet tồn tại (ExcelApi 1.1 không có `getItemOrNullObject`).
 * Ném lỗi liệt kê các sheet không tìm thấy (R2.9).
 */
export async function assertSheetsExist(context: Excel.RequestContext, sheetNames: readonly string[]): Promise<void> {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();
  const existing = new Set(worksheets.items.map((sheet) => sheet.name.toLowerCase()));
  const missing = Array.from(new Set(sheetNames)).filter((name) => !existing.has(name.toLowerCase()));
  if (missing.length > 0) {
    throw new UserFacingError(
      `Không tìm thấy sheet: ${missing.map((name) => `"${name}"`).join(", ")}. Sheet có thể đã bị xóa hoặc đổi tên.`,
    );
  }
}
