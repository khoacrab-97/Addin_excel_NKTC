import { columnIndexFromLetter, columnLetter, normalizeCellAddress } from "../address/address";

/** Tiêu đề bảng nhập tiến độ (F4) — trùng quy tắc nhận diện cột của F1 (R1.12). */
export const INPUT_TABLE_HEADERS = ["STT", "Nội dung công việc", "Bắt đầu", "Kết thúc"] as const;

/** Tên sheet mặc định khi tạo bảng trên sheet mới. */
export const INPUT_SHEET_BASE_NAME = "TienDo";
export const DEFAULT_TABLE_ROWS = 10;
export const MAX_TABLE_ROWS = 1_000;

const MAX_ROW = 1_048_576;
const MAX_COLUMN_INDEX = 16_383;

/** Vị trí các phần của bảng trên sheet (địa chỉ không gồm tên sheet). */
export interface InputTableLayout {
  /** Toàn bộ bảng: tiêu đề + dòng dữ liệu, ví dụ `B3:E13`. */
  address: string;
  header: string;
  /** Vùng dữ liệu (không gồm tiêu đề). */
  body: string;
  /** Cột STT / Nội dung / Bắt đầu / Kết thúc của vùng dữ liệu, ví dụ `B4:B13`. */
  columns: { stt: string; content: string; start: string; end: string };
  /** Ô đầu tiên cần nhập (Nội dung công việc của dòng 1). */
  firstInput: string;
  /** Giá trị cột STT: 1..n. */
  numbers: number[][];
  rows: number;
}

export type LayoutResult = { ok: true; layout: InputTableLayout } | { ok: false; error: string };

/** Đọc số dòng nhập từ hộp thoại: số nguyên 1..1000. */
export function parseRowCount(text: string): number | undefined {
  const trimmed = text.trim();
  if (!/^\d{1,4}$/.test(trimmed)) return undefined;
  const rows = Number(trimmed);
  return rows >= 1 && rows <= MAX_TABLE_ROWS ? rows : undefined;
}

/**
 * Tính vị trí bảng 4 cột từ ô bắt đầu (góc trên-trái, là ô tiêu đề "STT") và số dòng dữ liệu (R4.2, R4.3).
 */
export function buildInputTableLayout(startCell: string, rows: number): LayoutResult {
  const cell = normalizeCellAddress(startCell);
  if (cell === undefined) return { ok: false, error: `"${startCell}" không phải địa chỉ ô hợp lệ (ví dụ B3).` };
  if (!Number.isInteger(rows) || rows < 1 || rows > MAX_TABLE_ROWS) {
    return { ok: false, error: `Số dòng phải là số nguyên từ 1 đến ${MAX_TABLE_ROWS}.` };
  }
  const match = /^([A-Z]+)(\d+)$/.exec(cell);
  const firstColumn = columnIndexFromLetter(match?.[1] ?? "A");
  const headerRow = Number(match?.[2] ?? "1");
  const lastColumn = firstColumn + INPUT_TABLE_HEADERS.length - 1;
  const lastRow = headerRow + rows;
  if (lastColumn > MAX_COLUMN_INDEX || lastRow > MAX_ROW) {
    return { ok: false, error: "Bảng vượt quá giới hạn sheet của Excel. Chọn ô bắt đầu khác hoặc giảm số dòng." };
  }

  const letter = (offset: number): string => columnLetter(firstColumn + offset);
  const bodyColumn = (offset: number): string => `${letter(offset)}${headerRow + 1}:${letter(offset)}${lastRow}`;
  return {
    ok: true,
    layout: {
      address: `${letter(0)}${headerRow}:${letter(3)}${lastRow}`,
      header: `${letter(0)}${headerRow}:${letter(3)}${headerRow}`,
      body: `${letter(0)}${headerRow + 1}:${letter(3)}${lastRow}`,
      columns: { stt: bodyColumn(0), content: bodyColumn(1), start: bodyColumn(2), end: bodyColumn(3) },
      firstInput: `${letter(1)}${headerRow + 1}`,
      numbers: Array.from({ length: rows }, (_value, index) => [index + 1]),
      rows,
    },
  };
}

/** Ô đã có dữ liệu/công thức trong vùng sẽ tạo bảng (R4.4) — địa chỉ để báo cho người dùng. */
export function occupiedCells(
  values: readonly (readonly unknown[])[],
  formulas: readonly (readonly unknown[])[],
  topLeft: string,
): string[] {
  const match = /^([A-Z]+)(\d+)$/.exec(topLeft);
  const firstColumn = columnIndexFromLetter(match?.[1] ?? "A");
  const firstRow = Number(match?.[2] ?? "1");
  const occupied: string[] = [];
  values.forEach((row, r) => {
    row.forEach((value, c) => {
      const formula = formulas[r]?.[c];
      const hasFormula = typeof formula === "string" && formula.startsWith("=");
      if (hasFormula || (value !== "" && value !== null && value !== undefined)) {
        occupied.push(`${columnLetter(firstColumn + c)}${firstRow + r}`);
      }
    });
  });
  return occupied;
}
