import { matchHeaders } from "../mapping/headerMatch";
import type { ExportFormState } from "./protocol";

/**
 * Hộp thoại Xuất vào form: vùng nguồn đổi (chọn lại, thêm / chèn cột) → cột của từng dòng gán đang soạn và cột Ngày
 * được tìm lại theo tên trong dòng tiêu đề mới. Cột không còn → bỏ chọn cột của dòng đó và trả về thông báo.
 */
export function remapFormState(
  state: ExportFormState,
  previous: readonly string[],
  next: readonly string[],
): { state: ExportFormState; lost: string[] } {
  const positions = matchHeaders(previous, next);
  const lost: string[] = [];
  const rows = state.rows.map((row, index) => {
    if (row.column === undefined) return row;
    const column = positions[row.column];
    if (column !== undefined) return { ...row, column };
    lost.push(
      `Dòng ${String(index + 1)}: cột "${previous[row.column] ?? ""}" không còn trong vùng nguồn (bị xóa hoặc đổi tên) — chọn lại cột.`,
    );
    return { target: row.target };
  });
  const result: ExportFormState = { rows };
  if (state.selected !== undefined) result.selected = state.selected;
  if (state.dateColumn !== undefined) {
    const dateColumn = positions[state.dateColumn];
    if (dateColumn !== undefined) result.dateColumn = dateColumn;
    else lost.push(`Cột Ngày "${previous[state.dateColumn] ?? ""}" không còn trong vùng nguồn — chọn lại cột ngày.`);
  }
  return { state: result, lost };
}
