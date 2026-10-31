import { type DateSystem, type DayNumber, toDayNumber } from "../dates/excelDate";
import type { DailyLogColumns } from "./columns";

/** Một công việc hợp lệ trong danh sách input. */
export interface WorkItem {
  /** Số dòng trên sheet (1-based) để báo lỗi cho người dùng. */
  rowNumber: number;
  content: string;
  start: DayNumber;
  end: DayNumber;
}

/** Dòng input bị bỏ qua và lý do (R1.6). */
export interface RowIssue {
  rowNumber: number;
  reason: string;
}

export interface ParseWorkItemsResult {
  items: WorkItem[];
  issues: RowIssue[];
  /** Số dòng trống hoàn toàn ở 3 cột chính (bỏ qua, không tính là lỗi). */
  blankRowCount: number;
}

function isBlank(value: unknown): boolean {
  return value === undefined || value === null || (typeof value === "string" && value.trim() === "");
}

function toContentText(value: unknown): string {
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return "";
}

/**
 * Đọc danh sách công việc từ các dòng dữ liệu (không gồm dòng tiêu đề).
 * @param rows Giá trị ô (`range.values`) của các dòng dữ liệu.
 * @param firstRowNumber Số dòng trên sheet của phần tử `rows[0]`.
 */
export function parseWorkItems(
  rows: readonly (readonly unknown[])[],
  columns: DailyLogColumns,
  firstRowNumber: number,
  dateSystem: DateSystem = 1900,
): ParseWorkItemsResult {
  const items: WorkItem[] = [];
  const issues: RowIssue[] = [];
  let blankRowCount = 0;

  rows.forEach((row, index) => {
    const rowNumber = firstRowNumber + index;
    const rawContent = row[columns.content];
    const rawStart = row[columns.start];
    const rawEnd = row[columns.end];

    if (isBlank(rawContent) && isBlank(rawStart) && isBlank(rawEnd)) {
      blankRowCount++;
      return;
    }

    const content = toContentText(rawContent);
    if (content === "") {
      issues.push({ rowNumber, reason: "Thiếu nội dung công việc." });
      return;
    }

    const start = toDayNumber(rawStart, dateSystem);
    if (start === undefined) {
      issues.push({ rowNumber, reason: isBlank(rawStart) ? "Thiếu ngày bắt đầu." : "Ngày bắt đầu không hợp lệ." });
      return;
    }

    // R1.5: thiếu ngày kết thúc → công việc 1 ngày.
    const end = isBlank(rawEnd) ? start : toDayNumber(rawEnd, dateSystem);
    if (end === undefined) {
      issues.push({ rowNumber, reason: "Ngày kết thúc không hợp lệ." });
      return;
    }
    if (end < start) {
      issues.push({ rowNumber, reason: "Ngày kết thúc trước ngày bắt đầu." });
      return;
    }

    items.push({ rowNumber, content, start, end });
  });

  return { items, issues, blankRowCount };
}
