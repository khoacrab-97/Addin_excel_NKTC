/**
 * Giao thức giữa nút lệnh trên Ribbon (function file) và cửa sổ hộp thoại (dialog.html).
 * Office Dialog API 1.1 (Excel 2016): dữ liệu vào hộp thoại qua URL / localStorage, kết quả trả về qua
 * `messageParent(string)`. Hai phía đều kiểm tra dữ liệu nhận được (CLAUDE.md §3: dữ liệu bên ngoài).
 */

export type DialogView =
  "create-table" | "f1-columns" | "message" | "export-form" | "goto" | "confirm" | "help" | "about";

export const DIALOG_VIEWS: readonly DialogView[] = [
  "create-table",
  "f1-columns",
  "message",
  "export-form",
  "goto",
  "confirm",
  "help",
  "about",
];

/** Một cột của vùng dữ liệu để chọn trong hộp thoại, ví dụ `{ index: 2, label: "Cột C – Nội dung" }`. */
export interface ColumnOption {
  index: number;
  label: string;
}

export type MessageTone = "info" | "success" | "warning" | "error";

/** File tải về có sẵn trong add-in. */
export type DownloadKey = "form-template";

export interface CreateTablePayload {
  /** Sheet đang mở và ô đang chọn — gợi ý vị trí bảng trên sheet hiện tại. */
  sheetName: string;
  cell: string;
  defaultRows: number;
  maxRows: number;
}

/** Ô chọn vùng có nút mũi tên (giống ô chọn vùng của Excel): giá trị hiển thị + lỗi nếu vùng không dùng được. */
export interface RangeFieldState {
  /** Ví dụ `TienDo!A1:D11`. Chuỗi rỗng nếu chưa có vùng. */
  value: string;
  error?: string;
}

export interface F1ColumnsPayload {
  source: RangeFieldState;
  options: ColumnOption[];
  detected: { content?: number; start?: number; end?: number };
}

export interface MessagePayload {
  title: string;
  tone: MessageTone;
  lines: string[];
  /**
   * Liên kết tải về: chỉ là MÃ của file có sẵn trong add-in; hộp thoại tự đổi mã thành đường dẫn cố định
   * (không nhận đường dẫn từ dữ liệu truyền vào hộp thoại).
   */
  link?: DownloadKey;
  /** Bảng xem trước (ví dụ 5 ngày đầu của nhật ký vừa xuất). */
  table?: { headers: string[]; rows: string[][] };
}

/** Một dòng gán của hộp thoại "Xuất vào form": cột nguồn (trái) → ô trên form (phải). */
export interface FormRowState {
  column?: number;
  target: string;
}

/** Trạng thái người dùng đang nhập trong hộp thoại "Xuất vào form" (giữ lại khi mở lại hộp thoại). */
export interface ExportFormState {
  dateColumn?: number;
  rows: FormRowState[];
  /** Dòng đang chọn (để Xóa dòng). */
  selected?: number;
}

export interface ExportFormPayload {
  source: RangeFieldState;
  options: ColumnOption[];
  dayCount: number;
  state: ExportFormState;
  /** Lỗi của lần bấm Gán trước (theo dòng). */
  errors?: string[];
  /** Lỗi của ô đích cụ thể khi chọn bằng nút mũi tên. */
  rowError?: { row: number; message: string };
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
  "create-table": CreateTablePayload;
  "f1-columns": F1ColumnsPayload;
  message: MessagePayload;
  "export-form": ExportFormPayload;
  goto: GotoPayload;
  confirm: ConfirmPayload;
  help: Record<string, never>;
  about: AboutPayload;
}

export type CancelReply = { action: "cancel" };

export interface DialogReplies {
  "create-table": { action: "create"; target: "current" | "new"; cell: string; rows: number } | CancelReply;
  "f1-columns":
    | { action: "ok"; columns: { content: number; start: number; end: number } }
    /** Chọn lại vùng: `address` = địa chỉ gõ tay; không có → lấy vùng đang chọn trên sheet. */
    | { action: "pick"; address?: string }
    | CancelReply;
  message: { action: "ok" } | CancelReply;
  "export-form":
    | { action: "pick-source"; address?: string; state: ExportFormState }
    | { action: "pick-target"; row: number; address?: string; state: ExportFormState }
    | { action: "assign"; state: ExportFormState }
    | CancelReply;
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

function isPositiveCount(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 1_048_576;
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

function isRangeField(value: unknown): value is RangeFieldState {
  return isRecord(value) && isText(value.value) && isOptionalText(value.error);
}

function isFormRow(value: unknown): value is FormRowState {
  return isRecord(value) && isOptionalIndex(value.column) && isText(value.target) && value.target.length <= 80;
}

function isFormState(value: unknown): value is ExportFormState {
  return (
    isRecord(value) &&
    isOptionalIndex(value.dateColumn) &&
    isList(value.rows, isFormRow) &&
    value.rows.length <= 500 &&
    isOptionalIndex(value.selected)
  );
}

function isOptionalAddress(value: unknown): value is string | undefined {
  return value === undefined || (typeof value === "string" && value.length <= 80);
}

function isLink(value: unknown): boolean {
  return value === undefined || value === "form-template";
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
  "create-table": (p) =>
    isText(p.sheetName) && isText(p.cell) && isPositiveCount(p.defaultRows) && isPositiveCount(p.maxRows),
  "f1-columns": (p) =>
    isRangeField(p.source) &&
    isList(p.options, isColumnOption) &&
    isRecord(p.detected) &&
    isOptionalIndex(p.detected.content) &&
    isOptionalIndex(p.detected.start) &&
    isOptionalIndex(p.detected.end),
  message: (p) => isText(p.title) && isTone(p.tone) && isTextList(p.lines) && isTable(p.table) && isLink(p.link),
  "export-form": (p) =>
    isRangeField(p.source) &&
    isList(p.options, isColumnOption) &&
    typeof p.dayCount === "number" &&
    isFormState(p.state) &&
    (p.errors === undefined || isTextList(p.errors)) &&
    (p.rowError === undefined || (isRecord(p.rowError) && isIndex(p.rowError.row) && isText(p.rowError.message))),
  goto: (p) => isOptionalText(p.first) && isOptionalText(p.last) && isOptionalText(p.current),
  confirm: (p) => isText(p.day) && isList(p.cells, isConfirmCell),
  help: () => true,
  about: (p) => isText(p.platform) && isText(p.officeVersion) && isText(p.maxExcelApi),
};

const REPLY_GUARDS: Guards = {
  "create-table": (r) =>
    r.action === "create" &&
    (r.target === "current" || r.target === "new") &&
    isText(r.cell) &&
    r.cell.length <= 16 &&
    isPositiveCount(r.rows),
  "f1-columns": (r) =>
    (r.action === "ok" &&
      isRecord(r.columns) &&
      isIndex(r.columns.content) &&
      isIndex(r.columns.start) &&
      isIndex(r.columns.end)) ||
    (r.action === "pick" && isOptionalAddress(r.address)),
  message: (r) => r.action === "ok",
  "export-form": (r) =>
    isFormState(r.state) &&
    ((r.action === "pick-source" && isOptionalAddress(r.address)) ||
      (r.action === "pick-target" && isIndex(r.row) && isOptionalAddress(r.address)) ||
      r.action === "assign"),
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
