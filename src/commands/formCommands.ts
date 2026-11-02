import { formatReference, parseCellReference } from "../core/address/reference";
import { sameHeaders } from "../core/dailyLog/prepareDailyLog";
import { DAILY_SHEET_BASE_NAME } from "../core/dailyLog/sheetName";
import type { ColumnOption, ExportFormPayload, ExportFormState, RangeFieldState } from "../core/dialog/protocol";
import { columnOptions } from "../core/dialog/options";
import { headerTexts } from "../core/mapping/config";
import { buildDayIndex, detectDateColumn } from "../core/mapping/days";
import { blankHeaderColumns, buildMappingsFromRows } from "../core/mapping/formRows";
import { type MappingConfig, mappingKey } from "../core/mapping/types";
import { UserFacingError, toUserMessage } from "../excel/errors";
import {
  type RangeRef,
  type RangeSnapshot,
  getActiveSheetName,
  getSelectedCell,
  getSheetNames,
  getUsedRangeRef,
  headerHasMergedCells,
} from "../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../excel/settingsStore";
import { FORM_SHEET_NAME, canInsertTemplate, insertFormTemplate } from "../excel/templateInserter";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showMessage } from "./messages";
import { refText, resolveRangeRef, tryReadRange } from "./rangePick";
import { showDayCommand } from "./templateCommands";

const TITLE = "Xuất vào form";
const MAX_ROUNDS = 200;
const DAILY_SHEET = new RegExp(`^${DAILY_SHEET_BASE_NAME}(_\\d+)?$`, "i");

/** "Chèn form mẫu": thêm sheet "Mẫu NKTC" (biểu mẫu đã dựng sẵn) vào cuối workbook. */
export async function insertTemplateCommand(): Promise<void> {
  if (!canInsertTemplate()) {
    await showMessage(
      "Chèn form mẫu",
      "warning",
      [
        "Phiên bản Excel này chưa hỗ trợ chèn sheet từ file mẫu (cần Microsoft 365 hoặc Excel 2021 trở lên).",
        'Cách làm thay thế: tải biểu mẫu bên dưới → mở file → chuột phải tab sheet "Mẫu NKTC" → Move or Copy… → chọn file nhật ký, tích Create a copy → OK.',
      ],
      undefined,
      "form-template",
    );
    return;
  }
  const name = await insertFormTemplate();
  await showMessage("Chèn form mẫu", "success", [
    `Đã chèn sheet "${name}".`,
    'Tiếp theo: bấm "Xuất vào form" để chọn vùng nhật ký theo ngày và gán cột vào các ô của biểu mẫu.',
  ]);
}

/** Vùng nguồn mặc định: cấu hình đã lưu → sheet NK_TheoNgay đang mở hoặc mới nhất → vùng đang chọn. */
async function defaultSourceRef(config: MappingConfig | undefined): Promise<RangeRef> {
  if (config) return config.source;
  const [active, names] = await Promise.all([getActiveSheetName(), getSheetNames()]);
  const daily = DAILY_SHEET.test(active) ? active : names.filter((name) => DAILY_SHEET.test(name)).pop();
  return daily === undefined ? resolveRangeRef(undefined) : getUsedRangeRef(daily);
}

/** Sheet cho ô đích không ghi tên sheet: sheet của dòng gán đã có → "Mẫu NKTC" → sheet đang mở. */
function defaultFormSheet(state: ExportFormState, names: readonly string[], active: string): string {
  for (const row of state.rows) {
    const sheet = parseCellReference(row.target)?.sheetName;
    if (sheet !== undefined) return sheet;
  }
  return names.find((name) => name.toLowerCase() === FORM_SHEET_NAME.toLowerCase()) ?? active;
}

interface SourceCheck {
  field: RangeFieldState;
  snapshot?: RangeSnapshot;
  headers: string[];
}

/** Đọc vùng nguồn, dòng đầu là tiêu đề; dòng tiêu đề có ô gộp/trống → không dùng được (phải gỡ merge). */
async function checkSource(ref: RangeRef): Promise<SourceCheck> {
  const read = await tryReadRange(ref);
  if (!read.ok) return { field: { value: refText(ref), error: read.error }, headers: [] };
  const snapshot = read.snapshot;
  const value = refText(snapshot.ref);
  if (snapshot.values.length < 2) {
    return { field: { value, error: "Vùng nguồn cần có dòng tiêu đề và ít nhất một ngày." }, headers: [] };
  }
  const merged = await headerHasMergedCells(snapshot.ref);
  const headers = headerTexts(snapshot.values[0] ?? []);
  const blank = blankHeaderColumns(headers);
  if (merged === true || blank.length > 0) {
    const detail =
      merged === true
        ? "Dòng đầu tiên của vùng có ô bị gộp (merge)."
        : `Dòng đầu tiên có ô tiêu đề trống (cột thứ ${blank.map((column) => column + 1).join(", ")}) — thường do ô bị gộp (merge).`;
    return {
      field: {
        value,
        error: `${detail} Không chọn được vùng này: gỡ merge (Merge & Center → Unmerge Cells), đặt tên cho từng cột rồi chọn lại.`,
      },
      headers: [],
    };
  }
  return { field: { value }, snapshot, headers };
}

/** Giữ dòng gán hợp lệ với vùng nguồn mới (cột vượt số cột → bỏ chọn cột). */
function fitState(state: ExportFormState, columnCount: number): ExportFormState {
  const next: ExportFormState = {
    rows: state.rows.map((row) =>
      row.column !== undefined && row.column >= columnCount ? { target: row.target } : row,
    ),
  };
  if (state.selected !== undefined) next.selected = state.selected;
  if (state.dateColumn !== undefined && state.dateColumn < columnCount) next.dateColumn = state.dateColumn;
  return next;
}

/**
 * "Xuất vào form" (F2): một hộp thoại chung — vùng nguồn (ô chọn vùng có nút mũi tên, dòng đầu là tiêu đề),
 * các dòng gán "cột nguồn → ô trên form" (+ Thêm / Xóa dòng), nút Gán. Gán xong: lưu cấu hình, form nhận dữ liệu
 * ngày đầu tiên; các nút Chuyển ngày dùng cấu hình này.
 */
export async function exportFormCommand(): Promise<void> {
  const config = loadMappingConfig();
  const [names, active] = await Promise.all([getSheetNames(), getActiveSheetName()]);

  let ref: RangeRef | undefined;
  let pickError: string | undefined;
  try {
    ref = await defaultSourceRef(config);
  } catch (error) {
    pickError = toUserMessage(error);
  }
  let state: ExportFormState = {
    rows:
      config && config.mappings.length > 0
        ? config.mappings.map((mapping) => ({
            column: mapping.column,
            target: formatReference(mapping.target.sheetName, mapping.target.address),
          }))
        : [{ target: "" }],
    ...(config ? { dateColumn: config.dateColumn } : {}),
  };
  let errors: string[] | undefined;
  let rowError: ExportFormPayload["rowError"];

  for (let round = 0; round < MAX_ROUNDS; round++) {
    const source: SourceCheck = ref
      ? await checkSource(ref)
      : { field: { value: "", error: pickError ?? "Chưa có vùng nguồn." }, headers: [] };
    if (pickError !== undefined) source.field = { ...source.field, error: pickError };
    const snapshot = source.field.error === undefined ? source.snapshot : undefined;
    const options: ColumnOption[] = snapshot ? columnOptions(snapshot.values[0] ?? [], snapshot.firstColumnIndex) : [];
    if (snapshot) {
      state = fitState(state, source.headers.length);
      if (state.dateColumn === undefined) {
        const detected = detectDateColumn(snapshot.values);
        if (detected !== undefined) state = { ...state, dateColumn: detected };
      }
    }
    const dayCount =
      snapshot && state.dateColumn !== undefined
        ? buildDayIndex(snapshot.values, state.dateColumn, snapshot.firstRowNumber).days.length
        : 0;

    const reply = await openDialog(
      "export-form",
      {
        source: source.field,
        options,
        dayCount,
        state,
        ...(errors ? { errors } : {}),
        ...(rowError ? { rowError } : {}),
      },
      DIALOG_SIZE.large,
    );
    errors = undefined;
    rowError = undefined;
    if (reply.action === "cancel") return;
    state = reply.state;

    if (reply.action === "pick-source") {
      pickError = undefined;
      try {
        ref = await resolveRangeRef(reply.address);
      } catch (error) {
        pickError = toUserMessage(error);
      }
      continue;
    }

    if (reply.action === "pick-target") {
      const row = state.rows[reply.row];
      if (!row) continue;
      try {
        let text: string;
        if (reply.address === undefined) {
          const cell = await getSelectedCell();
          text = formatReference(cell.sheetName, cell.address);
        } else {
          const cell = parseCellReference(reply.address);
          if (!cell) throw new UserFacingError(`"${reply.address}" không phải một ô hợp lệ (ví dụ Mẫu NKTC!E2).`);
          text = formatReference(cell.sheetName ?? defaultFormSheet(state, names, active), cell.address);
        }
        state = {
          ...state,
          rows: state.rows.map((item, index) => (index === reply.row ? { ...item, target: text } : item)),
        };
      } catch (error) {
        rowError = { row: reply.row, message: toUserMessage(error) };
      }
      continue;
    }

    // Gán
    if (!snapshot) {
      errors = ["Vùng nguồn chưa hợp lệ — chọn lại vùng nguồn."];
      continue;
    }
    if (state.dateColumn === undefined) {
      errors = ["Chưa chọn cột Ngày (dùng để chuyển ngày)."];
      continue;
    }
    const built = buildMappingsFromRows(state.rows, source.headers.length, defaultFormSheet(state, names, active));
    if (!built.ok) {
      errors = built.errors;
      continue;
    }
    const sheets = new Set((await getSheetNames()).map((name) => name.toLowerCase()));
    const missing = Array.from(new Set(built.mappings.map((mapping) => mapping.target.sheetName))).filter(
      (name) => !sheets.has(name.toLowerCase()),
    );
    if (missing.length > 0) {
      errors = [`Không có sheet ${missing.map((name) => `"${name}"`).join(", ")} trong file.`];
      continue;
    }
    if (buildDayIndex(snapshot.values, state.dateColumn, snapshot.firstRowNumber).days.length === 0) {
      errors = ["Cột Ngày đã chọn không có ngày hợp lệ."];
      continue;
    }

    const keepConfirmation =
      config !== undefined &&
      config.confirmedKey !== undefined &&
      config.confirmedKey === mappingKey(built.mappings) &&
      sameHeaders(source.headers, config.headers);
    await saveMappingConfig({
      version: 1,
      source: snapshot.ref,
      headers: source.headers,
      dateColumn: state.dateColumn,
      mappings: built.mappings,
      ...(keepConfirmation && config.confirmedKey !== undefined ? { confirmedKey: config.confirmedKey } : {}),
    });
    const day = await showDayCommand({ kind: "first" });
    if (day !== undefined) {
      await showMessage(TITLE, "success", [
        `Đã gán ${built.mappings.length} ô; biểu mẫu đang hiển thị dữ liệu ngày đầu tiên.`,
        "Dùng các nút Ngày đầu / Ngày trước / Ngày sau / Ngày cuối / Đến ngày… để chuyển ngày.",
      ]);
    }
    return;
  }
}
