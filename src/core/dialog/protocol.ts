/**
 * Giao thức giữa nút lệnh trên Ribbon (function file) và cửa sổ hộp thoại (dialog.html).
 * Office Dialog API 1.1 (Excel 2016): dữ liệu vào hộp thoại qua URL / localStorage, kết quả trả về qua
 * `messageParent(string)`. Hai phía đều kiểm tra dữ liệu nhận được (CLAUDE.md §3: dữ liệu bên ngoài).
 */

export type DialogView = "f1-columns" | "message" | "assign" | "settings" | "goto" | "confirm" | "help" | "about";

export const DIALOG_VIEWS: readonly DialogView[] = [
  "f1-columns",
  "message",
  "assign",
  "settings",
  "goto",
  "confirm",
  "help",
  "about",
];

/** Một cột của vùng dữ liệu để chọn trong hộp thoại, ví dụ `{ index: 2, label: "C – Nội dung" }`. */
export interface ColumnOption {
  index: number;
  label: string;
}

export type MessageTone = "info" | "success" | "warning" | "error";

export interface F1ColumnsPayload {
  source: string;
  options: ColumnOption[];
  detected: { content?: number; start?: number; end?: number };
}

export interface MessagePayload {
  title: string;
  tone: MessageTone;
  lines: string[];
  /** Bảng xem trước (ví dụ 5 ngày đầu của nhật ký vừa xuất). */
  table?: { headers: string[]; rows: string[][] };
}

export interface AssignPayload {
  target: string;
  source: string;
  options: ColumnOption[];
  dateColumn: number;
  /** Cột đang gán cho ô này (nếu có). */
  current?: number;
}

export interface MappingRow {
  key: string;
  target: string;
  column: number;
}

export interface SettingsPayload {
  source: string;
  options: ColumnOption[];
  dateColumn: number;
  mappings: MappingRow[];
  dayCount: number;
  currentDay?: string;
}

export interface GotoPayload {
  first?: string;
  last?: string;
  current?: string;
}

export interface ConfirmCell {
  target: string;
  formula: boolean;
  value: boolean;
}

export interface ConfirmPayload {
  day: string;
  cells: ConfirmCell[];
}

export interface AboutPayload {
  platform: string;
  officeVersion: string;
  maxExcelApi: string;
}

export interface DialogPayloads {
  "f1-columns": F1ColumnsPayload;
  message: MessagePayload;
  assign: AssignPayload;
  settings: SettingsPayload;
  goto: GotoPayload;
  confirm: ConfirmPayload;
  help: Record<string, never>;
  about: AboutPayload;
}

export type CancelReply = { action: "cancel" };

export interface DialogReplies {
  "f1-columns": { action: "ok"; columns: { content: number; start: number; end: number } } | CancelReply;
  message: { action: "ok" } | CancelReply;
  assign: { action: "assign"; column: number } | { action: "unassign" } | CancelReply;
  settings: { action: "save"; dateColumn: number; remove: string[] } | CancelReply;
  goto: { action: "goto"; text: string } | CancelReply;
  confirm: { action: "ok" } | CancelReply;
  help: { action: "ok" } | CancelReply;
  about: { action: "ok" } | CancelReply;
}

/** Giới hạn độ dài nhãn cột gửi sang hộp thoại (URL của Internet Explorer 11 tối đa ~2.000 ký tự). */
export const LABEL_MAX_LENGTH = 60;
const LIST_MAX = 2_000;
const TEXT_MAX = 2_000;

// ---------- Kiểm tra kiểu dữ liệu ----------

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isIndex(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value < 16_384;
}

function isText(value: unknown): value is string {
  return typeof value === "string" && value.length <= TEXT_MAX;
}

function isOptionalText(value: unknown): value is string | undefined {
  return value === undefined || isText(value);
}

function isList<T>(value: unknown, guard: (item: unknown) => item is T): value is T[] {
  return Array.isArray(value) && value.length <= LIST_MAX && value.every(guard);
}

function isTextList(value: unknown): value is string[] {
  return isList(value, isText);
}

function isColumnOption(value: unknown): value is ColumnOption {
  return isRecord(value) && isIndex(value.index) && isText(value.label);
}

function isOptionalIndex(value: unknown): value is number | undefined {
  return value === undefined || isIndex(value);
}

function isTone(value: unknown): value is MessageTone {
  return value === "info" || value === "success" || value === "warning" || value === "error";
}

function isMappingRow(value: unknown): value is MappingRow {
  return isRecord(value) && isText(value.key) && isText(value.target) && isIndex(value.column);
}

function isConfirmCell(value: unknown): value is ConfirmCell {
  return (
    isRecord(value) && isText(value.target) && typeof value.formula === "boolean" && typeof value.value === "boolean"
  );
}

function isTable(value: unknown): value is MessagePayload["table"] {
  return (
    value === undefined ||
    (isRecord(value) && isTextList(value.headers) && isList(value.rows, (row): row is string[] => isTextList(row)))
  );
}

type Guards = { [V in DialogView]: (value: Record<string, unknown>) => boolean };

const PAYLOAD_GUARDS: Guards = {
  "f1-columns": (p) =>
    isText(p.source) &&
    isList(p.options, isColumnOption) &&
    isRecord(p.detected) &&
    isOptionalIndex(p.detected.content) &&
    isOptionalIndex(p.detected.start) &&
    isOptionalIndex(p.detected.end),
  message: (p) => isText(p.title) && isTone(p.tone) && isTextList(p.lines) && isTable(p.table),
  assign: (p) =>
    isText(p.target) &&
    isText(p.source) &&
    isList(p.options, isColumnOption) &&
    isIndex(p.dateColumn) &&
    isOptionalIndex(p.current),
  settings: (p) =>
    isText(p.source) &&
    isList(p.options, isColumnOption) &&
    isIndex(p.dateColumn) &&
    isList(p.mappings, isMappingRow) &&
    typeof p.dayCount === "number" &&
    isOptionalText(p.currentDay),
  goto: (p) => isOptionalText(p.first) && isOptionalText(p.last) && isOptionalText(p.current),
  confirm: (p) => isText(p.day) && isList(p.cells, isConfirmCell),
  help: () => true,
  about: (p) => isText(p.platform) && isText(p.officeVersion) && isText(p.maxExcelApi),
};

const REPLY_GUARDS: Guards = {
  "f1-columns": (r) =>
    r.action === "ok" &&
    isRecord(r.columns) &&
    isIndex(r.columns.content) &&
    isIndex(r.columns.start) &&
    isIndex(r.columns.end),
  message: (r) => r.action === "ok",
  assign: (r) => (r.action === "assign" && isIndex(r.column)) || r.action === "unassign",
  settings: (r) => r.action === "save" && isIndex(r.dateColumn) && isTextList(r.remove),
  goto: (r) => r.action === "goto" && isText(r.text),
  confirm: (r) => r.action === "ok",
  help: (r) => r.action === "ok",
  about: (r) => r.action === "ok",
};

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return undefined;
  }
}

export function isDialogView(value: unknown): value is DialogView {
  return typeof value === "string" && (DIALOG_VIEWS as readonly string[]).includes(value);
}

/** Đọc dữ liệu đầu vào của hộp thoại; `undefined` nếu không hợp lệ. */
export function parsePayload<V extends DialogView>(view: V, text: string): DialogPayloads[V] | undefined {
  const value = parseJson(text);
  return isRecord(value) && PAYLOAD_GUARDS[view](value) ? (value as unknown as DialogPayloads[V]) : undefined;
}

/** Đọc kết quả hộp thoại gửi về; dữ liệu không hợp lệ được coi như người dùng hủy. */
export function parseReply<V extends DialogView>(view: V, text: string): DialogReplies[V] {
  const value = parseJson(text);
  if (isRecord(value) && REPLY_GUARDS[view](value)) return value as unknown as DialogReplies[V];
  return { action: "cancel" };
}

/** Rút gọn nhãn cho vừa URL hộp thoại. */
export function shortLabel(text: string): string {
  const trimmed = text.replace(/\s+/g, " ").trim();
  return trimmed.length <= LABEL_MAX_LENGTH ? trimmed : `${trimmed.slice(0, LABEL_MAX_LENGTH - 1)}…`;
}
