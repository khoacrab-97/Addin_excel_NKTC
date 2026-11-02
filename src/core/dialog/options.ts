import { columnLetter } from "../address/address";
import { type ColumnOption, shortLabel } from "./protocol";

/** Danh sách cột để chọn trong hộp thoại: "C – Nội dung công việc". */
export function columnOptions(headers: readonly unknown[], firstColumnIndex: number): ColumnOption[] {
  return headers.map((header, index) => {
    const text =
      typeof header === "string"
        ? header
        : typeof header === "number" || typeof header === "boolean"
          ? String(header)
          : "";
    const label = shortLabel(text);
    return {
      index,
      label: `${columnLetter(firstColumnIndex + index)} – ${label === "" ? "(không có tiêu đề)" : label}`,
    };
  });
}
