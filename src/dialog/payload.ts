import { type DialogPayloads, type DialogView, isDialogView, parsePayload } from "../core/dialog/protocol";

/** Tiền tố khóa localStorage — trùng với `PAYLOAD_STORAGE_PREFIX` của function file. */
const PAYLOAD_STORAGE_PREFIX = "nktc.dialog.";

/** Đọc một tham số trong query string (Internet Explorer 11 không có URLSearchParams). */
export function queryParam(search: string, name: string): string | undefined {
  for (const part of search.replace(/^\?/, "").split("&")) {
    const [key, value] = part.split("=");
    if (key === name && value !== undefined) {
      try {
        return decodeURIComponent(value);
      } catch {
        return undefined;
      }
    }
  }
  return undefined;
}

function readStored(id: string): string | undefined {
  try {
    return window.localStorage.getItem(PAYLOAD_STORAGE_PREFIX + id) ?? undefined;
  } catch {
    return undefined;
  }
}

function readHash(): string | undefined {
  const hash = window.location.hash.replace(/^#/, "");
  if (hash === "") return undefined;
  try {
    return decodeURIComponent(hash);
  } catch {
    return undefined;
  }
}

export type LoadedDialog = { [V in DialogView]: { view: V; payload: DialogPayloads[V] } }[DialogView];

/** Xác định hộp thoại cần hiển thị và dữ liệu đầu vào (URL trước, localStorage sau). */
export function loadDialog(): LoadedDialog | undefined {
  const view = queryParam(window.location.search, "view");
  const id = queryParam(window.location.search, "id") ?? "";
  if (!isDialogView(view)) return undefined;
  for (const text of [readHash(), readStored(id)]) {
    if (text === undefined) continue;
    const payload = parsePayload(view, text);
    if (payload !== undefined) return { view, payload } as LoadedDialog;
  }
  return undefined;
}
