import { detectColumns } from "../core/dailyLog/columns";
import { prepareDailyLog, sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { formatDay } from "../core/dates/excelDate";
import type { ColumnOption, RangeFieldState } from "../core/dialog/protocol";
import { columnOptions } from "../core/dialog/options";
import { writeDailyLogSheet } from "../excel/dailyLogWriter";
import { toUserMessage } from "../excel/errors";
import { type RangeRef, readRange } from "../excel/rangeReader";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";
import { type SnapshotResult, refText, resolveRangeRef, tryReadRange } from "./rangePick";

const TITLE = "Xuất nhật ký theo ngày";
const MAX_ISSUES_SHOWN = 10;
const PREVIEW_DAYS = 5;
const PREVIEW_TEXT_LENGTH = 300;
const MAX_PICKS = 50;

function truncate(text: string, length: number): string {
  return text.length <= length ? text : `${text.slice(0, length - 1)}…`;
}

function fieldState(ref: RangeRef | undefined, read: SnapshotResult | undefined, error?: string): RangeFieldState {
  const value = ref ? refText(ref) : "";
  if (error !== undefined) return { value, error };
  if (!read) return { value, error: "Chưa có vùng dữ liệu." };
  if (!read.ok) return { value, error: read.error };
  if (read.snapshot.values.length < 2) {
    return { value: refText(read.snapshot.ref), error: "Vùng cần có dòng tiêu đề và ít nhất một dòng công việc." };
  }
  return { value: refText(read.snapshot.ref) };
}

/**
 * F1 từ Ribbon: hộp thoại có ô "Vùng dữ liệu" (tự nhận diện từ vùng/ô đang chọn; nút mũi tên để lấy lại vùng
 * đang chọn trên sheet, hoặc gõ địa chỉ) và 3 ô chọn cột; đọc lại dữ liệu mới nhất rồi tạo sheet NK_TheoNgay.
 */
export async function exportDailyLogCommand(): Promise<void> {
  let ref: RangeRef | undefined;
  let pickError: string | undefined;
  try {
    ref = await resolveRangeRef(undefined);
  } catch (error) {
    pickError = toUserMessage(error);
  }
  let read = ref ? await tryReadRange(ref) : undefined;

  for (let pick = 0; pick < MAX_PICKS; pick++) {
    const source = fieldState(ref, read, pickError);
    const snapshot = read?.ok === true && source.error === undefined ? read.snapshot : undefined;
    const headers = snapshot?.values[0] ?? [];
    const options: ColumnOption[] = snapshot ? columnOptions(headers, snapshot.firstColumnIndex) : [];
    const reply = await openDialog(
      "f1-columns",
      { source, options, detected: snapshot ? detectColumns(headers) : {} },
      DIALOG_SIZE.medium,
    );
    if (reply.action === "cancel") return;
    if (reply.action === "pick") {
      pickError = undefined;
      try {
        ref = await resolveRangeRef(reply.address);
        read = await tryReadRange(ref);
      } catch (error) {
        pickError = toUserMessage(error);
      }
      continue;
    }
    if (!snapshot) return;
    await exportFrom(snapshot.ref, headers, snapshot.firstColumnIndex, reply.columns);
    return;
  }
}

async function exportFrom(
  ref: RangeRef,
  headers: readonly unknown[],
  firstColumnIndex: number,
  columns: { content: number; start: number; end: number },
): Promise<void> {
  // Đọc lại dữ liệu mới nhất trước khi xuất (CLAUDE.md §4).
  const fresh = await readRange(ref);
  if (!sameHeaders(fresh.values[0] ?? [], headers) || fresh.firstColumnIndex !== firstColumnIndex) {
    await showMessage(TITLE, "warning", [
      "Dòng tiêu đề của vùng dữ liệu đã thay đổi trong lúc chọn cột. Vui lòng bấm Xuất theo ngày lần nữa.",
    ]);
    return;
  }

  const result = prepareDailyLog(fresh.values, columns, fresh.firstRowNumber);
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
  lines.push('Có thể thêm cột (Thời tiết, Số công nhân…) vào sheet này, rồi dùng "Xuất vào form".');
  await showMessage(TITLE, issues.length > 0 ? "warning" : "success", lines, {
    headers: ["Ngày", "Nội dung"],
    rows: result.built.entries
      .slice(0, PREVIEW_DAYS)
      .map((entry) => [formatDay(entry.day), truncate(entry.content, PREVIEW_TEXT_LENGTH)]),
  });
}
