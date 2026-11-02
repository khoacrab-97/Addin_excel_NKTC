import { DEFAULT_TABLE_ROWS, MAX_TABLE_ROWS, buildInputTableLayout } from "../core/inputTable/layout";
import { UserFacingError } from "../excel/errors";
import { createInputTable, getActiveCell } from "../excel/inputTableWriter";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";

/**
 * F4 "Tạo bảng": tạo bảng nhập tiến độ STT | Nội dung công việc | Bắt đầu | Kết thúc để người dùng tự điền,
 * sau đó dùng "Xuất theo ngày" (F1) trên bảng này.
 */
export async function createTableCommand(): Promise<void> {
  const active = await getActiveCell();
  const reply = await openDialog(
    "create-table",
    { sheetName: active.sheetName, cell: active.cell, defaultRows: DEFAULT_TABLE_ROWS, maxRows: MAX_TABLE_ROWS },
    DIALOG_SIZE.medium,
  );
  if (reply.action !== "create") return;

  // Bảng trên sheet mới luôn bắt đầu ở A1.
  const result = buildInputTableLayout(reply.target === "new" ? "A1" : reply.cell, reply.rows);
  if (!result.ok) throw new UserFacingError(result.error);
  const created = await createInputTable({ kind: reply.target }, result.layout);
  await showMessage("Tạo bảng", "success", [
    `Đã tạo bảng ${created.sheetName}!${created.address} gồm ${result.layout.rows} dòng.`,
    "Điền Nội dung công việc, ngày Bắt đầu, ngày Kết thúc (dd/mm/yyyy). Gõ tiếp ngay dưới dòng cuối, bảng tự mở rộng.",
    'Điền xong: đặt con trỏ trong bảng → bấm "Xuất theo ngày".',
  ]);
}
