import { parseRangeReference } from "../core/address/reference";
import { UserFacingError, toUserMessage } from "../excel/errors";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import {
  type RangeRef,
  type RangeSnapshot,
  getActiveSheetName,
  getSelectedTableRef,
  readRange,
} from "../excel/rangeReader";

/**
 * Vùng cho ô chọn vùng (nút mũi tên): không có `address` → vùng đang chọn trên sheet (chọn 1 ô → cả vùng có
 * dữ liệu của sheet); có `address` (người dùng gõ) → đọc địa chỉ, không ghi tên sheet thì là sheet đang mở.
 */
export async function resolveRangeRef(address: string | undefined): Promise<RangeRef> {
  if (address === undefined) return getSelectedTableRef();
  const reference = parseRangeReference(address);
  if (!reference) throw new UserFacingError(`"${address}" không phải địa chỉ vùng hợp lệ (ví dụ TienDo!A1:D11).`);
  return { sheetName: reference.sheetName ?? (await getActiveSheetName()), address: reference.address };
}

export type SnapshotResult = { ok: true; snapshot: RangeSnapshot } | { ok: false; error: string };

/** Đọc vùng cho hộp thoại; lỗi (sheet không còn, vùng trống…) trả về thông điệp để hiện ngay dưới ô chọn vùng. */
export async function tryReadRange(ref: RangeRef): Promise<SnapshotResult> {
  try {
    return { ok: true, snapshot: await readRange(ref) };
  } catch (error) {
    return { ok: false, error: toUserMessage(error) };
  }
}

export function refText(ref: RangeRef): string {
  return `${ref.sheetName}!${ref.address}`;
}

/**
 * Nút ⬆ của ô chọn vùng: hộp thoại chính đã đóng (ẩn); mở hộp thoại nhỏ "chọn trên sheet rồi bấm OK".
 * OK → `true` (bên gọi đọc vùng/ô đang chọn); Hủy/đóng → `false` (giữ nguyên giá trị cũ).
 */
export async function pickOnSheet(kind: "range" | "cell", label: string): Promise<boolean> {
  const reply = await openDialog("pick", { kind, label }, DIALOG_SIZE.pick);
  return reply.action === "ok";
}
