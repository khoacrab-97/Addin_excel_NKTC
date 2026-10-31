import type { CellWrite } from "../core/mapping/writePlan";
import { type TargetCell, formatTarget } from "../core/mapping/types";
import { UserFacingError } from "./errors";
import { assertSheetsExist } from "./worksheets";

/** Trạng thái hiện tại của một ô đích trên form mẫu. */
export interface TargetState {
  target: TargetCell;
  hasFormula: boolean;
  hasValue: boolean;
}

function isInvalidAddressError(error: unknown): boolean {
  if (typeof error !== "object" || error === null || !("code" in error)) return false;
  return error.code === "InvalidArgument" || error.code === "InvalidReference";
}

/** Đọc trạng thái các ô đích (công thức / dữ liệu sẵn có) để xin xác nhận trước khi ghi (R2.4, R2.5). */
export async function inspectTargets(targets: readonly TargetCell[]): Promise<TargetState[]> {
  return Excel.run(async (context) => {
    await assertSheetsExist(
      context,
      targets.map((target) => target.sheetName),
    );
    const ranges = targets.map((target) => {
      const range = context.workbook.worksheets.getItem(target.sheetName).getRange(target.address);
      range.load(["formulas", "values"]);
      return range;
    });
    try {
      await context.sync();
    } catch (error) {
      if (isInvalidAddressError(error)) throw new UserFacingError("Có ô đích không hợp lệ trong cấu hình.");
      throw error;
    }
    return targets.map((target, index) => {
      const range = ranges[index];
      const formula: unknown = range?.formulas[0]?.[0];
      const value: unknown = range?.values[0]?.[0];
      return {
        target,
        hasFormula: typeof formula === "string" && formula.startsWith("="),
        hasValue: value !== "" && value !== null && value !== undefined,
      };
    });
  });
}

/**
 * Ghi giá trị vào các ô đích trong một lần đồng bộ (R2.3, R2.6): chỉ gán `values`,
 * không đổi định dạng/merge/viền. Kiểm tra sheet trước, thiếu sheet thì không ghi ô nào (R2.9).
 */
export async function writeCells(writes: readonly CellWrite[]): Promise<void> {
  if (writes.length === 0) return;
  await Excel.run(async (context) => {
    await assertSheetsExist(
      context,
      writes.map((write) => write.target.sheetName),
    );
    for (const write of writes) {
      context.workbook.worksheets.getItem(write.target.sheetName).getRange(write.target.address).values = [
        [write.value],
      ];
    }
    try {
      await context.sync();
    } catch (error) {
      if (isInvalidAddressError(error)) {
        throw new UserFacingError(
          `Không ghi được vào ô đích (${writes.map((w) => formatTarget(w.target)).join(", ")}).`,
        );
      }
      throw error;
    }
  });
}
