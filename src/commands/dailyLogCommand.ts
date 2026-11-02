import { detectColumns } from "../core/dailyLog/columns";
import { prepareDailyLog, sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { formatDay } from "../core/dates/excelDate";
import { columnOptions } from "../core/dialog/options";
import { writeDailyLogSheet } from "../excel/dailyLogWriter";
import { getSelectedTableRef, readRange } from "../excel/rangeReader";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";

const TITLE = "Xuất nhật ký theo ngày";
const MAX_ISSUES_SHOWN = 10;
const PREVIEW_DAYS = 5;
const PREVIEW_TEXT_LENGTH = 300;

function truncate(text: string, length: number): string {
  return text.length <= length ? text : `${text.slice(0, length - 1)}…`;
}

/**
 * F1 từ Ribbon: lấy bảng tiến độ đang chọn (hoặc cả vùng dữ liệu của sheet nếu chỉ chọn 1 ô),
 * xác nhận 3 cột trong hộp thoại, đọc lại dữ liệu mới nhất rồi tạo sheet NK_TheoNgay (R1.x).
 */
export async function exportDailyLogCommand(): Promise<void> {
  const ref = await getSelectedTableRef();
  const snapshot = await readRange(ref);
  const headers = snapshot.values[0] ?? [];
  if (snapshot.values.length < 2) {
    await showMessage(TITLE, "warning", [
      "Vùng dữ liệu cần có dòng tiêu đề và ít nhất một dòng công việc.",
      "Bôi chọn bảng tiến độ (gồm dòng tiêu đề) hoặc đặt con trỏ trong sheet tiến độ rồi bấm lại.",
    ]);
    return;
  }

  const reply = await openDialog(
    "f1-columns",
    {
      source: `${snapshot.ref.sheetName}!${snapshot.ref.address}`,
      options: columnOptions(headers, snapshot.firstColumnIndex),
      detected: detectColumns(headers),
    },
    DIALOG_SIZE.medium,
  );
  if (reply.action !== "ok") return;

  // Đọc lại dữ liệu mới nhất trước khi xuất (CLAUDE.md §4).
  const fresh = await readRange(snapshot.ref);
  if (!sameHeaders(fresh.values[0] ?? [], headers) || fresh.firstColumnIndex !== snapshot.firstColumnIndex) {
    await showMessage(TITLE, "warning", [
      "Dòng tiêu đề của vùng dữ liệu đã thay đổi trong lúc chọn cột. Vui lòng bấm Xuất theo ngày lần nữa.",
    ]);
    return;
  }

  const result = prepareDailyLog(fresh.values, reply.columns, fresh.firstRowNumber);
  const issues = result.parsed?.issues ?? [];
  const issueLines = issues.slice(0, MAX_ISSUES_SHOWN).map((issue) => `Dòng ${issue.rowNumber}: ${issue.reason}`);
  if (issues.length > MAX_ISSUES_SHOWN) issueLines.push(`… và ${issues.length - MAX_ISSUES_SHOWN} dòng khác.`);

  if (!result.ok) {
    await showMessage(TITLE, "error", [result.error, ...issueLines]);
    return;
  }

  const sheetName = await writeDailyLogSheet(result.built.entries);
  const lines = [
    `Đã tạo sheet "${sheetName}" gồm ${result.built.entries.length} ngày, từ ${formatDay(result.built.firstDay)} đến ${formatDay(result.built.lastDay)}.`,
    `${result.parsed.items.length} công việc hợp lệ.`,
  ];
  if (issues.length > 0) lines.push(`${issues.length} dòng lỗi đã bỏ qua:`, ...issueLines);
  await showMessage(TITLE, issues.length > 0 ? "warning" : "success", lines, {
    headers: ["Ngày", "Nội dung"],
    rows: result.built.entries
      .slice(0, PREVIEW_DAYS)
      .map((entry) => [formatDay(entry.day), truncate(entry.content, PREVIEW_TEXT_LENGTH)]),
  });
}
