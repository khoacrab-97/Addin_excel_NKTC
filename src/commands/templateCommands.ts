import { sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { formatDay, parseDateText } from "../core/dates/excelDate";
import { columnOptions } from "../core/dialog/options";
import { headerTexts } from "../core/mapping/config";
import { buildDayIndex, detectDateColumn } from "../core/mapping/days";
import { type DayCommand, resolveDay } from "../core/mapping/navigation";
import { type MappingConfig, formatTarget, mappingKey, targetKey } from "../core/mapping/types";
import { validateMappings } from "../core/mapping/validate";
import { buildWritePlan } from "../core/mapping/writePlan";
import { UserFacingError } from "../excel/errors";
import { type RangeSnapshot, getSelectedCell, getSelectedTableRef, readRange } from "../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../excel/settingsStore";
import { inspectTargets, writeCells } from "../excel/templateWriter";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";

const TITLE = "Form mẫu";
const HEADERS_CHANGED =
  "Dòng tiêu đề của vùng nguồn đã thay đổi so với cấu hình. Chọn lại vùng nguồn (nút Vùng nguồn) rồi kiểm tra gán cột.";

function requireConfig(): MappingConfig {
  const config = loadMappingConfig();
  if (!config) {
    throw new UserFacingError(
      'Chưa có cấu hình form mẫu. Bôi chọn bảng nhật ký theo ngày (gồm tiêu đề) rồi bấm "Vùng nguồn" trước.',
    );
  }
  return config;
}

function sourceText(config: MappingConfig): string {
  return `${config.source.sheetName}!${config.source.address}`;
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

/** "Vùng nguồn": lấy bảng nhật ký theo ngày đang chọn làm nguồn của form mẫu (R2.1, R2.8). */
export async function setSourceCommand(): Promise<void> {
  const fresh = await readRange(await getSelectedTableRef());
  if (fresh.values.length < 2) {
    await showMessage(TITLE, "warning", [
      "Vùng nguồn cần có dòng tiêu đề và ít nhất một ngày.",
      "Bôi chọn bảng nhật ký theo ngày (ví dụ sheet NK_TheoNgay) hoặc đặt con trỏ trong sheet đó rồi bấm lại.",
    ]);
    return;
  }
  const previous = loadMappingConfig();
  const headers = headerTexts(fresh.values[0] ?? []);
  const keep = previous !== undefined && sameHeaders(headers, previous.headers);
  const detected = detectDateColumn(fresh.values);
  const next: MappingConfig = keep
    ? { ...previous, source: fresh.ref, headers }
    : { version: 1, source: fresh.ref, headers, dateColumn: detected ?? 0, mappings: [] };
  await saveMappingConfig(next);

  const options = columnOptions(fresh.values[0] ?? [], fresh.firstColumnIndex);
  const index = buildDayIndex(fresh.values, next.dateColumn, fresh.firstRowNumber);
  const lines = [
    `Vùng nguồn: ${sourceText(next)}.`,
    `Cột Ngày: ${options[next.dateColumn]?.label ?? "?"}${!keep && detected === undefined ? " (chưa nhận diện được — kiểm tra trong Danh sách gán)" : ""}.`,
    `${index.days.length} ngày hợp lệ${index.issues.length > 0 ? `, ${index.issues.length} dòng ngày lỗi bị bỏ qua` : ""}.`,
  ];
  if (keep) lines.push(`Giữ nguyên ${next.mappings.length} ô đã gán.`);
  else if (previous && previous.mappings.length > 0) lines.push("Tiêu đề khác cấu hình cũ → đã bỏ các ô gán cũ.");
  lines.push('Tiếp theo: chọn ô trên biểu mẫu → bấm "Gán cột" (lặp lại cho từng ô) → dùng các nút chuyển ngày.');
  await showMessage(TITLE, index.days.length > 0 ? "success" : "warning", lines);
}

/** "Gán cột": gán một cột của vùng nguồn cho ô (hoặc ô merge) đang chọn trên biểu mẫu (R2.1, R2.2). */
export async function assignCellCommand(): Promise<void> {
  const config = requireConfig();
  const target = await getSelectedCell();
  const fresh = await readSource(config);
  if (!fresh) return;
  const owner = config.mappings.find((mapping) => targetKey(mapping.target) === targetKey(target));
  const reply = await openDialog(
    "assign",
    {
      target: formatTarget(target),
      source: sourceText(config),
      options: columnOptions(fresh.values[0] ?? [], fresh.firstColumnIndex),
      dateColumn: config.dateColumn,
      ...(owner ? { current: owner.column } : {}),
    },
    DIALOG_SIZE.medium,
  );
  if (reply.action === "cancel") return;

  const others = config.mappings.filter((mapping) => targetKey(mapping.target) !== targetKey(target));
  if (reply.action === "unassign") {
    await saveMappingConfig({ ...config, mappings: others });
    return;
  }
  if (reply.column >= config.headers.length) throw new UserFacingError("Cột đã chọn không nằm trong vùng nguồn.");
  await saveMappingConfig({ ...config, mappings: [...others, { column: reply.column, target }] });
}

/** "Danh sách gán": xem/xóa các ô đã gán, đổi cột Ngày. */
export async function mappingListCommand(): Promise<void> {
  const config = requireConfig();
  const fresh = await readSource(config);
  if (!fresh) return;
  const options = columnOptions(fresh.values[0] ?? [], fresh.firstColumnIndex);
  const index = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
  const reply = await openDialog(
    "settings",
    {
      source: sourceText(config),
      options,
      dateColumn: config.dateColumn,
      mappings: config.mappings.map((mapping) => ({
        key: targetKey(mapping.target),
        target: formatTarget(mapping.target),
        column: mapping.column,
      })),
      dayCount: index.days.length,
      ...(config.currentDay !== undefined ? { currentDay: formatDay(config.currentDay) } : {}),
    },
    DIALOG_SIZE.large,
  );
  if (reply.action !== "save") return;
  if (reply.dateColumn >= config.headers.length) throw new UserFacingError("Cột Ngày không nằm trong vùng nguồn.");
  const removed = new Set(reply.remove);
  await saveMappingConfig({
    ...config,
    dateColumn: reply.dateColumn,
    mappings: config.mappings.filter((mapping) => !removed.has(targetKey(mapping.target))),
  });
}

/**
 * Ghi dữ liệu một ngày lên form (R2.3–R2.7): đọc lại nguồn, chọn ngày, xin xác nhận nếu ô đích đang có
 * công thức/dữ liệu và bộ ô đích chưa được xác nhận, rồi ghi trong một lần đồng bộ.
 */
export async function showDayCommand(command: DayCommand): Promise<void> {
  const config = requireConfig();
  const fresh = await readSource(config);
  if (!fresh) return;
  const errors = validateMappings(config.mappings, config.headers.length);
  if (errors.length > 0) {
    await showMessage(TITLE, "warning", [...errors, 'Chọn ô trên biểu mẫu rồi bấm "Gán cột" để gán.']);
    return;
  }
  const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
  const resolution = resolveDay(days, config.currentDay, command);
  if (!resolution.ok) {
    await showMessage(TITLE, resolution.severity === "info" ? "info" : "error", [resolution.message]);
    return;
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
      if (reply.action !== "ok") return;
    }
  }

  await writeCells(buildWritePlan(fresh.values[resolution.target.valueRowIndex] ?? [], config.mappings));
  await saveMappingConfig({ ...config, currentDay: resolution.target.day, confirmedKey: key });
  if (resolution.notice) await showMessage(TITLE, "warning", [resolution.notice]);
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
