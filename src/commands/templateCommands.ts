import { sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { isDatePattern, weekdayName } from "../core/dates/datePattern";
import { type DayNumber, formatDay, parseDateText } from "../core/dates/excelDate";
import type { ConfirmPayload, DayPanelPayload, MessageTone } from "../core/dialog/protocol";
import { headerTexts } from "../core/mapping/config";
import { type SourceDay, buildDayIndex, indexOfDay } from "../core/mapping/days";
import { type DayCommand, resolveDay } from "../core/mapping/navigation";
import { type MappingConfig, formatTarget, mappingKey, targetKey } from "../core/mapping/types";
import { validateMappings } from "../core/mapping/validate";
import { buildWritePlan } from "../core/mapping/writePlan";
import { UserFacingError, toUserMessage } from "../excel/errors";
import { type RangeSnapshot, readRange } from "../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../excel/settingsStore";
import { inspectTargets, writeCells } from "../excel/templateWriter";
import { DIALOG_SIZE, openDialog, openPersistentDialog } from "./dialogHost";
import { showMessage } from "./messages";

const TITLE = "Chuyển ngày";
const HEADERS_CHANGED =
  'Dòng tiêu đề của vùng nguồn đã thay đổi so với cấu hình. Bấm "Xuất vào form" để chọn lại vùng nguồn và kiểm tra gán cột.';
const NOT_ASSIGNED =
  'Chưa gán cột vào form. Bấm "Xuất vào form" (nhóm Form mẫu) để chọn vùng nguồn và gán cột trước — chuyển ngày chỉ hoạt động sau khi đã gán.';

function requireConfig(): MappingConfig {
  const config = loadMappingConfig();
  if (!config || config.mappings.length === 0) throw new UserFacingError(NOT_ASSIGNED);
  return config;
}

/** Đọc lại vùng nguồn (R2.7) và kiểm tra tiêu đề còn khớp cấu hình. */
async function readSource(config: MappingConfig): Promise<RangeSnapshot> {
  const fresh = await readRange(config.source);
  if (!sameHeaders(headerTexts(fresh.values[0] ?? []), config.headers)) throw new UserFacingError(HEADERS_CHANGED);
  return fresh;
}

/** Kết quả một lần chuyển ngày. */
export type DayOutcome =
  | { kind: "shown"; day: DayNumber; notice?: string }
  | { kind: "message"; tone: MessageTone; text: string }
  | { kind: "cancelled" };

/** Hỏi người dùng có ghi đè các ô đích đang có dữ liệu/công thức không (R2.5). */
export type ConfirmOverwrite = (payload: ConfirmPayload) => Promise<boolean>;

async function confirmByDialog(payload: ConfirmPayload): Promise<boolean> {
  const reply = await openDialog("confirm", payload, DIALOG_SIZE.medium);
  return reply.action === "ok";
}

/**
 * Ghi dữ liệu một ngày lên form (R2.3–R2.7, R2.15): đọc lại nguồn, chọn ngày, xin xác nhận nếu ô đích đang có
 * công thức/dữ liệu và bộ ô đích chưa được xác nhận (ô mẫu "Ngày … tháng … năm …" của cột Ngày không tính),
 * rồi ghi trong một lần đồng bộ.
 */
export async function navigateDay(
  command: DayCommand,
  confirm: ConfirmOverwrite = confirmByDialog,
): Promise<DayOutcome> {
  const config = requireConfig();
  const fresh = await readSource(config);
  const errors = validateMappings(config.mappings, config.headers.length);
  if (errors.length > 0) {
    return { kind: "message", tone: "warning", text: `${errors.join(" ")} Bấm "Xuất vào form" để gán cột.` };
  }
  const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
  const resolution = resolveDay(days, config.currentDay, command);
  if (!resolution.ok) {
    return { kind: "message", tone: resolution.severity === "info" ? "info" : "error", text: resolution.message };
  }

  const states = await inspectTargets(config.mappings.map((mapping) => mapping.target));
  const texts = new Map(states.map((state) => [targetKey(state.target), state.text]));
  const dateTargets = new Set(
    config.mappings.filter((mapping) => mapping.column === config.dateColumn).map((m) => targetKey(m.target)),
  );
  const key = mappingKey(config.mappings);
  if (config.confirmedKey !== key) {
    const risky = states.filter((state) => {
      const datePattern = dateTargets.has(targetKey(state.target)) && isDatePattern(state.text ?? "");
      return state.hasFormula || (state.hasValue && !datePattern);
    });
    if (risky.length > 0) {
      const ok = await confirm({
        day: formatDay(resolution.target.day),
        cells: risky.map((state) => ({
          target: formatTarget(state.target),
          formula: state.hasFormula,
          value: state.hasValue,
        })),
      });
      if (!ok) return { kind: "cancelled" };
    }
  }

  const row = fresh.values[resolution.target.valueRowIndex] ?? [];
  await writeCells(
    buildWritePlan(row, config.mappings, {
      dateColumn: config.dateColumn,
      currentText: (target) => texts.get(targetKey(target)),
    }),
  );
  await saveMappingConfig({ ...config, currentDay: resolution.target.day, confirmedKey: key });
  return resolution.notice === undefined
    ? { kind: "shown", day: resolution.target.day }
    : { kind: "shown", day: resolution.target.day, notice: resolution.notice };
}

/** Nút chuyển ngày trên Ribbon: ghi ngày; thông báo (đầu/cuối, lỗi) hiện bằng hộp thoại nhỏ. */
export async function showDayCommand(command: DayCommand): Promise<DayNumber | undefined> {
  const outcome = await navigateDay(command);
  if (outcome.kind === "message") await showMessage(TITLE, outcome.tone, [outcome.text]);
  if (outcome.kind === "shown") {
    if (outcome.notice !== undefined) await showMessage(TITLE, "warning", [outcome.notice]);
    return outcome.day;
  }
  return undefined;
}

function dayLabel(day: DayNumber): string {
  return `${weekdayName(day)}, ${formatDay(day)}`;
}

/** Trạng thái hiển thị của bảng chuyển ngày: ngày đang xem, vị trí, khoảng ngày. */
async function panelPayload(message?: DayPanelPayload["message"]): Promise<DayPanelPayload> {
  const payload: DayPanelPayload = {};
  if (message) payload.message = message;
  try {
    const config = requireConfig();
    const fresh = await readSource(config);
    const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
    const first: SourceDay | undefined = days[0];
    const last: SourceDay | undefined = days[days.length - 1];
    if (first && last) payload.range = { first: formatDay(first.day), last: formatDay(last.day), count: days.length };
    if (config.currentDay !== undefined) {
      const position = indexOfDay(days, config.currentDay);
      payload.current = { label: dayLabel(config.currentDay), value: formatDay(config.currentDay) };
      if (position !== -1) payload.current.position = position + 1;
    }
  } catch (error) {
    payload.message ??= { tone: "warning", text: toUserMessage(error) };
    payload.disabled = true;
  }
  return payload;
}

function isClosed(session: { closed: boolean }): boolean {
  return session.closed;
}

/** Hàm đóng bảng chuyển ngày đang mở (để nút Ribbon khác được chạy ngay). */
let closeActivePanel: (() => void) | undefined;

export function interruptDayPanel(): boolean {
  if (!closeActivePanel) return false;
  closeActivePanel();
  return true;
}

/**
 * "Bảng chuyển ngày": cửa sổ nhỏ nổi, để mở trong lúc làm việc — hiện ngày đang xem (thứ, ngày, vị trí),
 * ô gõ ngày + Enter, nút Đầu / Trước / Sau / Cuối. DialogApi 1.2 cập nhật tại chỗ; bản cũ mở lại cửa sổ.
 */
export async function dayPanelCommand(): Promise<void> {
  let panel = await openPersistentDialog("day-panel", await panelPayload(), DIALOG_SIZE.panel);
  // Đối tượng (không phải biến boolean) vì giá trị được đổi trong hàm đóng bảng gọi từ nút Ribbon khác.
  const session = { closed: false };
  closeActivePanel = () => {
    session.closed = true;
    panel.close();
  };
  try {
    for (;;) {
      const reply = await panel.next();
      if (session.closed) return;
      if (reply.action === "cancel") {
        // Nút Đóng của bảng gửi "cancel" nhưng hộp thoại mở liên tục không tự đóng → add-in đóng.
        panel.close();
        return;
      }

      let command: DayCommand | undefined;
      let message: DayPanelPayload["message"];
      if (reply.action === "goto") {
        const day = parseDateText(reply.text);
        if (day === undefined)
          message = { tone: "error", text: `"${reply.text}" không phải ngày hợp lệ (dd/mm/yyyy).` };
        else command = { kind: "goto", day };
      } else {
        command = { kind: reply.kind };
      }

      if (command) {
        try {
          const outcome = await navigateDay(command, async (payload) => {
            // Chỉ mở được một hộp thoại: tạm đóng bảng để hỏi xác nhận, sau đó mở lại.
            panel.close();
            const ok = await confirmByDialog(payload);
            panel = await openPersistentDialog("day-panel", await panelPayload(), DIALOG_SIZE.panel);
            return ok;
          });
          if (outcome.kind === "message") message = { tone: outcome.tone, text: outcome.text };
          if (outcome.kind === "shown" && outcome.notice !== undefined)
            message = { tone: "warning", text: outcome.notice };
          if (outcome.kind === "cancelled") message = { tone: "info", text: "Đã hủy, biểu mẫu giữ nguyên." };
        } catch (error) {
          message = { tone: "error", text: toUserMessage(error) };
        }
      }
      // Có thể vừa bị đóng bởi nút Ribbon khác trong lúc đang ghi (closeActivePanel).
      if (isClosed(session)) return;
      const next = await panelPayload(message);
      if (!panel.update(next)) {
        panel.close();
        panel = await openPersistentDialog("day-panel", next, DIALOG_SIZE.panel);
      }
    }
  } finally {
    closeActivePanel = undefined;
  }
}
