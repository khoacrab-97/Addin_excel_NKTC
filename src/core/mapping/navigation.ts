import { type DayNumber, formatDay } from "../dates/excelDate";
import { type SourceDay, indexOfDay } from "./days";

/** Lệnh điều hướng ngày từ Ribbon (F2). */
export type DayCommand =
  { kind: "first" } | { kind: "prev" } | { kind: "next" } | { kind: "last" } | { kind: "goto"; day: DayNumber };

export type DayResolution =
  /** Ngày cần ghi lên form; `notice` là thông báo kèm theo (không chặn việc ghi). */
  | { ok: true; target: SourceDay; notice?: string }
  /** Không ghi gì: `severity` "info" khi đã ở đầu/cuối danh sách, "error" khi ngày không hợp lệ. */
  | { ok: false; severity: "info" | "error"; message: string };

function range(days: readonly SourceDay[]): string {
  const first = days[0];
  const last = days[days.length - 1];
  return first && last ? `từ ${formatDay(first.day)} đến ${formatDay(last.day)}` : "";
}

/**
 * Xác định ngày cần hiển thị cho một lệnh điều hướng (R2 — điều hướng ngày).
 * `current` là ngày đang hiển thị (lưu trong cấu hình); nếu không còn trong vùng nguồn,
 * "Trước"/"Sau" quay về ngày đầu tiên kèm thông báo.
 */
export function resolveDay(
  days: readonly SourceDay[],
  current: DayNumber | undefined,
  command: DayCommand,
): DayResolution {
  const first = days[0];
  const last = days[days.length - 1];
  if (!first || !last) {
    return { ok: false, severity: "error", message: "Vùng nguồn không có ngày hợp lệ ở cột Ngày đã chọn." };
  }

  switch (command.kind) {
    case "first":
      return { ok: true, target: first };
    case "last":
      return { ok: true, target: last };
    case "goto": {
      const index = indexOfDay(days, command.day);
      const found = days[index];
      if (!found) {
        return {
          ok: false,
          severity: "error",
          message: `Ngày ${formatDay(command.day)} không có trong vùng nguồn (${range(days)}).`,
        };
      }
      return { ok: true, target: found };
    }
    case "prev":
    case "next": {
      const index = current === undefined ? -1 : indexOfDay(days, current);
      if (index === -1) {
        const notice =
          current === undefined
            ? undefined
            : `Không còn ngày ${formatDay(current)} trong vùng nguồn; đã hiển thị ngày ${formatDay(first.day)}.`;
        return notice === undefined ? { ok: true, target: first } : { ok: true, target: first, notice };
      }
      const nextIndex = command.kind === "next" ? index + 1 : index - 1;
      const target = days[nextIndex];
      if (!target) {
        return {
          ok: false,
          severity: "info",
          message:
            command.kind === "next"
              ? `Đang ở ngày cuối cùng (${formatDay(last.day)}).`
              : `Đang ở ngày đầu tiên (${formatDay(first.day)}).`,
        };
      }
      return { ok: true, target };
    }
  }
}
