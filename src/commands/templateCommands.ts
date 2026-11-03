import { isDatePattern } from "../core/dates/datePattern";
import { type DayNumber, formatDay } from "../core/dates/excelDate";
import type { ConfirmPayload, MessageTone } from "../core/dialog/protocol";
import { headerTexts } from "../core/mapping/config";
import { type MissingColumn, remapConfig } from "../core/mapping/headerMatch";
import { buildDayIndex } from "../core/mapping/days";
import { type DayCommand, resolveDay } from "../core/mapping/navigation";
import { type MappingConfig, formatTarget, mappingKey, targetKey } from "../core/mapping/types";
import { validateMappings } from "../core/mapping/validate";
import { buildWritePlan } from "../core/mapping/writePlan";
import { UserFacingError } from "../excel/errors";
import { type RangeSnapshot, expandSourceRef, readRange } from "../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../excel/settingsStore";
import { inspectTargets, writeCells } from "../excel/templateWriter";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";
import { refText } from "./rangePick";

const TITLE = "Chuyển ngày";
const NOT_ASSIGNED =
  'Chưa gán cột vào form. Bấm "Xuất vào form" (nhóm Form mẫu) để chọn vùng nguồn và gán cột trước — chuyển ngày chỉ hoạt động sau khi đã gán.';

function requireConfig(): MappingConfig {
  const config = loadMappingConfig();
  if (!config || config.mappings.length === 0) throw new UserFacingError(NOT_ASSIGNED);
  return config;
}

function missingMessage(missing: readonly MissingColumn[], source: string): string {
  const items = missing.map((item) =>
    item.target === undefined
      ? `cột Ngày "${item.header}"`
      : `cột "${item.header}" (gán vào ${formatTarget(item.target)})`,
  );
  return (
    `Không tìm thấy ${items.join(", ")} trong dòng tiêu đề của vùng nguồn ${source} — cột có thể đã bị xóa hoặc ` +
    'đổi tên tiêu đề. Biểu mẫu chưa được ghi. Bấm "Xuất vào form" để chọn lại cột cho các dòng này.'
  );
}

/**
 * Đọc lại vùng nguồn (R2.7) — tự mở rộng tới hết dữ liệu của sheet — và khớp lại cột đã gán theo tên tiêu đề
 * (R2.17): thêm / chèn / đổi chỗ cột vẫn ghi đúng; cột bị xóa / đổi tên → báo, không ghi.
 */
async function readSource(config: MappingConfig): Promise<{ fresh: RangeSnapshot; config: MappingConfig }> {
  const fresh = await readRange(await expandSourceRef(config.source));
  const remapped = remapConfig(config, headerTexts(fresh.values[0] ?? []));
  if (!remapped.ok) throw new UserFacingError(missingMessage(remapped.missing, refText(fresh.ref)));
  return { fresh, config: remapped.config };
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
  const { fresh, config } = await readSource(requireConfig());
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
