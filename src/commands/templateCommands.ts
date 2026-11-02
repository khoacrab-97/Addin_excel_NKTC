import { sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { formatDay, parseDateText } from "../core/dates/excelDate";
import { headerTexts } from "../core/mapping/config";
import { buildDayIndex } from "../core/mapping/days";
import { type DayCommand, resolveDay } from "../core/mapping/navigation";
import { type MappingConfig, formatTarget, mappingKey } from "../core/mapping/types";
import { validateMappings } from "../core/mapping/validate";
import { buildWritePlan } from "../core/mapping/writePlan";
import { UserFacingError } from "../excel/errors";
import { type RangeSnapshot, readRange } from "../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../excel/settingsStore";
import { inspectTargets, writeCells } from "../excel/templateWriter";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";

const TITLE = "Form mẫu";
const HEADERS_CHANGED =
  'Dòng tiêu đề của vùng nguồn đã thay đổi so với cấu hình. Bấm "Xuất vào form" để chọn lại vùng nguồn và kiểm tra gán cột.';

function requireConfig(): MappingConfig {
  const config = loadMappingConfig();
  if (!config || config.mappings.length === 0) {
    throw new UserFacingError(
      'Chưa gán cột vào form. Bấm "Xuất vào form" (nhóm Form mẫu) để chọn vùng nguồn và gán cột trước — các nút Chuyển ngày chỉ hoạt động sau khi đã gán.',
    );
  }
  return config;
}

/** Đọc lại vùng nguồn (R2.7) và kiểm tra tiêu đề còn khớp cấu hình. */
async function readSource(config: MappingConfig): Promise<RangeSnapshot | undefined> {
  const fresh = await readRange(config.source);
  if (!sameHeaders(headerTexts(fresh.values[0] ?? []), config.headers)) {
    await showMessage(TITLE, "warning", [HEADERS_CHANGED]);
    return undefined;
  }
  return fresh;
}

/**
 * Ghi dữ liệu một ngày lên form (R2.3–R2.7): đọc lại nguồn, chọn ngày, xin xác nhận nếu ô đích đang có
 * công thức/dữ liệu và bộ ô đích chưa được xác nhận, rồi ghi trong một lần đồng bộ.
 */
export async function showDayCommand(command: DayCommand): Promise<number | undefined> {
  const config = requireConfig();
  const fresh = await readSource(config);
  if (!fresh) return undefined;
  const errors = validateMappings(config.mappings, config.headers.length);
  if (errors.length > 0) {
    await showMessage(TITLE, "warning", [...errors, 'Bấm "Xuất vào form" để gán cột.']);
    return undefined;
  }
  const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
  const resolution = resolveDay(days, config.currentDay, command);
  if (!resolution.ok) {
    await showMessage(TITLE, resolution.severity === "info" ? "info" : "error", [resolution.message]);
    return undefined;
  }

  const key = mappingKey(config.mappings);
  if (config.confirmedKey !== key) {
    const states = await inspectTargets(config.mappings.map((mapping) => mapping.target));
    const risky = states.filter((state) => state.hasFormula || state.hasValue);
    if (risky.length > 0) {
      const reply = await openDialog(
        "confirm",
        {
          day: formatDay(resolution.target.day),
          cells: risky.map((state) => ({
            target: formatTarget(state.target),
            formula: state.hasFormula,
            value: state.hasValue,
          })),
        },
        DIALOG_SIZE.medium,
      );
      if (reply.action !== "ok") return undefined;
    }
  }

  await writeCells(buildWritePlan(fresh.values[resolution.target.valueRowIndex] ?? [], config.mappings));
  await saveMappingConfig({ ...config, currentDay: resolution.target.day, confirmedKey: key });
  if (resolution.notice) await showMessage(TITLE, "warning", [resolution.notice]);
  return resolution.target.day;
}

/** "Đến ngày…": nhập ngày dd/mm/yyyy rồi hiển thị ngày đó. */
export async function goToDayCommand(): Promise<void> {
  const config = requireConfig();
  const fresh = await readSource(config);
  if (!fresh) return;
  const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
  const first = days[0];
  const last = days[days.length - 1];
  const reply = await openDialog(
    "goto",
    {
      ...(first ? { first: formatDay(first.day) } : {}),
      ...(last ? { last: formatDay(last.day) } : {}),
      ...(config.currentDay !== undefined ? { current: formatDay(config.currentDay) } : {}),
    },
    DIALOG_SIZE.small,
  );
  if (reply.action !== "goto") return;
  const day = parseDateText(reply.text);
  if (day === undefined) {
    await showMessage(TITLE, "error", [`"${reply.text}" không phải ngày hợp lệ (định dạng dd/mm/yyyy).`]);
    return;
  }
  await showDayCommand({ kind: "goto", day });
}
