import { formatReference, parseRangeReference } from "../core/address/reference";
import { UserFacingError, toUserMessage } from "../excel/errors";
import { DIALOG_SIZE, openPersistentDialog } from "./dialogHost";
import {
  type RangeRef,
  type RangeSnapshot,
  getActiveSheetName,
  getSelectedCell,
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

/** Vùng/ô đang chọn trên sheet, dạng hiển thị `Sheet!A1:D20` (vùng: chọn 1 ô → cả vùng có dữ liệu của sheet). */
async function selectionText(kind: "range" | "cell"): Promise<string> {
  if (kind === "range") return refText(await getSelectedTableRef());
  const cell = await getSelectedCell();
  return formatReference(cell.sheetName, cell.address);
}

/** Đăng ký nghe "người dùng chọn lại trên sheet"; trả về hàm hủy đăng ký, `undefined` nếu không đăng ký được. */
function onSelectionChanged(handler: () => void): Promise<(() => void) | undefined> {
  const document = Office.context.document as Partial<Office.Document> | undefined;
  if (!document?.addHandlerAsync) return Promise.resolve(undefined);
  return new Promise((resolve) => {
    document.addHandlerAsync?.(Office.EventType.DocumentSelectionChanged, handler, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(undefined);
        return;
      }
      resolve(() => {
        document.removeHandlerAsync?.(Office.EventType.DocumentSelectionChanged, { handler });
      });
    });
  });
}

/**
 * Nút mũi tên của ô chọn vùng — giống Excel (ví dụ New Formatting Rule): hộp thoại chính đã đóng; mở thanh nhỏ
 * gồm tiêu đề, ô nhập và nút mũi tên. Người dùng chọn trên sheet → địa chỉ hiện ngay trong ô nhập; bấm mũi tên
 * (hoặc Enter) → trả về địa chỉ đang hiện trong ô nhập; đóng thanh (X) → `undefined` (giữ giá trị cũ).
 * Excel không cập nhật được ô nhập (Excel 2016/2019, DialogApi 1.1): bấm mũi tên → lấy vùng/ô đang chọn
 * trên sheet lúc đó (trừ khi người dùng tự gõ địa chỉ).
 */
export async function pickOnSheet(kind: "range" | "cell", title: string, current: string): Promise<string | undefined> {
  const payload = { kind, title, value: current };
  const bar = await openPersistentDialog("pick", payload, DIALOG_SIZE.pick);
  const stop = bar.canUpdate
    ? await onSelectionChanged(() => {
        selectionText(kind)
          .then((value) => {
            bar.update({ ...payload, value });
          })
          .catch(() => {
            // Chọn nhiều vùng rời nhau… → giữ địa chỉ đang hiện.
          });
      })
    : undefined;
  try {
    const reply = await bar.next();
    bar.close();
    if (reply.action !== "ok") return undefined;
    const text = reply.text.trim();
    if (reply.typed || stop !== undefined) return text === "" ? undefined : text;
    return await selectionText(kind);
  } finally {
    stop?.();
  }
}
