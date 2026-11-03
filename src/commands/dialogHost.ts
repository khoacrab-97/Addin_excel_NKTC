import {
  CURRENT_DIALOG_KEY,
  type DialogPayloads,
  type DialogReplies,
  type DialogView,
  NO_PAYLOAD_MESSAGE,
  formatDialogPointer,
  parseDialogPointer,
  parseReply,
} from "../core/dialog/protocol";
import { UserFacingError } from "../excel/errors";

/** Kích thước hộp thoại (% màn hình). */
export interface DialogSize {
  width: number;
  height: number;
}

export const DIALOG_SIZE = {
  small: { width: 30, height: 40 },
  medium: { width: 40, height: 60 },
  large: { width: 50, height: 75 },
  /** Thanh chọn trên sheet (hộp thoại thu gọn): tiêu đề + ô nhập + nút mũi tên. */
  pick: { width: 30, height: 20 },
} as const satisfies Record<string, DialogSize>;

/** Tiền tố khóa localStorage chứa dữ liệu đầu vào của hộp thoại. */
export const PAYLOAD_STORAGE_PREFIX = "nktc.dialog.";
/** URL hộp thoại trên Internet Explorer 11 (Excel 2016) tối đa ~2.000 ký tự. */
const MAX_URL_LENGTH = 1_900;

/** Lưu dữ liệu hộp thoại và con trỏ "hộp thoại đang mở"; `true` → mở được bằng đường dẫn ngắn. */
function storePayload(view: DialogView, id: string, json: string): boolean {
  try {
    window.localStorage.setItem(PAYLOAD_STORAGE_PREFIX + id, json);
    window.localStorage.setItem(CURRENT_DIALOG_KEY, formatDialogPointer(view, id));
    return true;
  } catch {
    // localStorage bị chặn: hộp thoại vẫn đọc được dữ liệu từ URL nếu đủ ngắn.
    return false;
  }
}

function removePayload(id: string): void {
  try {
    window.localStorage.removeItem(PAYLOAD_STORAGE_PREFIX + id);
    if (parseDialogPointer(window.localStorage.getItem(CURRENT_DIALOG_KEY))?.id === id) {
      window.localStorage.removeItem(CURRENT_DIALOG_KEY);
    }
  } catch {
    // Không ảnh hưởng chức năng.
  }
}

/** Origin của add-in (Internet Explorer 11 không có `location.origin`). */
function addinOrigin(): string {
  return `${window.location.protocol}//${window.location.host}`;
}

/** URL thư mục chứa function file (dialog.html nằm cùng thư mục, cùng domain — yêu cầu của Dialog API). */
function baseUrl(): string {
  const href = window.location.href;
  const cut = [href.indexOf("?"), href.indexOf("#")].filter((index) => index !== -1);
  const path = cut.length > 0 ? href.slice(0, Math.min(...cut)) : href;
  return path.slice(0, path.lastIndexOf("/") + 1);
}

function dialogErrorMessage(code: number): string {
  switch (code) {
    case 12007:
      return "Đang có một hộp thoại khác của add-in mở. Vui lòng đóng hộp thoại đó rồi thử lại.";
    case 12004:
    case 12005:
      return "Không mở được hộp thoại của add-in (địa chỉ không hợp lệ). Vui lòng cài lại add-in.";
    case 12009:
      return "Hộp thoại của add-in bị chặn. Vui lòng cho phép cửa sổ bật lên rồi thử lại.";
    default:
      return `Không mở được hộp thoại của add-in (mã lỗi ${code}).`;
  }
}

/** Hộp thoại vừa đóng: Excel cần một lúc mới đóng hẳn, mở hộp thoại mới ngay sẽ báo 12007 → chờ rồi thử lại. */
const BUSY_CODE = 12007;
const RETRY_DELAY_MS = 150;
const RETRY_LIMIT = 20;

/** `displayDialogAsync`, tự thử lại khi Excel báo đang có hộp thoại khác (tối đa ~3 giây). */
function displayDialog(
  url: string,
  size: DialogSize,
  callback: (result: Office.AsyncResult<Office.Dialog>) => void,
  attempt = 0,
): void {
  Office.context.ui.displayDialogAsync(
    url,
    { width: size.width, height: size.height, displayInIframe: true },
    (result) => {
      if (
        result.status === Office.AsyncResultStatus.Failed &&
        result.error.code === BUSY_CODE &&
        attempt < RETRY_LIMIT
      ) {
        window.setTimeout(() => {
          displayDialog(url, size, callback, attempt + 1);
        }, RETRY_DELAY_MS);
        return;
      }
      callback(result);
    },
  );
}

function newId(): string {
  return `${String(Date.now())}${Math.random().toString(36).slice(2, 8)}`;
}

/** Đường dẫn đầy đủ: tên hộp thoại + mã dữ liệu trong query, dữ liệu trong hash (nếu đủ ngắn). */
function fullDialogUrl(view: DialogView, json: string, id: string): string {
  let url = `${baseUrl()}dialog.html?view=${view}&id=${id}`;
  const hash = encodeURIComponent(json);
  if (url.length + hash.length + 1 <= MAX_URL_LENGTH) url += `#${hash}`;
  return url;
}

/** Cửa sổ hộp thoại hiện tại (đổi khi phải mở lại bằng đường dẫn đầy đủ). */
interface DialogWindow {
  current?: Office.Dialog;
}

/**
 * Mở cửa sổ hộp thoại. Office ghi địa chỉ trang lên thanh tiêu đề (không đổi được) → mở bằng đường dẫn ngắn
 * `…/dialog.html`, dữ liệu qua localStorage. Hộp thoại không đọc được dữ liệu (`NO_PAYLOAD_MESSAGE`) → đóng và
 * mở lại một lần bằng đường dẫn đầy đủ (cách cũ). `onMessage` chỉ nhận tin từ trang hộp thoại của add-in.
 */
function openWindow(
  view: DialogView,
  json: string,
  id: string,
  size: DialogSize,
  onMessage: (message: string, dialog: Office.Dialog) => void,
  onClosed: () => void,
): Promise<DialogWindow> {
  const short = storePayload(view, id, json);
  const win: DialogWindow = {};
  let fellBack = !short;
  return new Promise((resolve, reject) => {
    const show = (url: string, first: boolean): void => {
      displayDialog(url, size, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          if (first) {
            removePayload(id);
            reject(new UserFacingError(dialogErrorMessage(result.error.code)));
          } else {
            onClosed();
          }
          return;
        }
        const dialog = result.value;
        win.current = dialog;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          if (!("message" in args)) return;
          // Chỉ nhận kết quả từ trang hộp thoại của add-in. `origin` không có trên bản Office cũ
          // (DialogOrigin 1.1) — khi đó Office chỉ chuyển tin từ trang cùng domain đã mở.
          if (args.origin !== undefined && args.origin !== addinOrigin()) return;
          if (args.message === NO_PAYLOAD_MESSAGE) {
            if (fellBack) return;
            fellBack = true;
            dialog.close();
            show(fullDialogUrl(view, json, id), false);
            return;
          }
          onMessage(args.message, dialog);
        });
        // Người dùng đóng hộp thoại (12006) hoặc trang hộp thoại lỗi.
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (win.current === dialog) onClosed();
        });
        if (first) resolve(win);
      });
    };
    show(short ? `${baseUrl()}dialog.html` : fullDialogUrl(view, json, id), true);
  });
}

/**
 * Mở hộp thoại (Office Dialog API 1.1) và chờ kết quả. Người dùng đóng hộp thoại bằng nút X → `{ action: "cancel" }`.
 * Dữ liệu đầu vào được gửi qua localStorage (cùng domain với hộp thoại), dự phòng qua URL.
 */
export function openDialog<V extends DialogView>(
  view: V,
  payload: DialogPayloads[V],
  size: DialogSize = DIALOG_SIZE.medium,
): Promise<DialogReplies[V]> {
  const json = JSON.stringify(payload);
  const id = newId();
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (reply: DialogReplies[V]): void => {
      if (settled) return;
      settled = true;
      removePayload(id);
      resolve(reply);
    };
    openWindow(
      view,
      json,
      id,
      size,
      (message, dialog) => {
        dialog.close();
        finish(parseReply(view, message));
      },
      () => {
        finish({ action: "cancel" });
      },
    ).catch(reject);
  });
}

/** Hộp thoại mở liên tục (thanh chọn trên sheet): nhận nhiều tin, cập nhật nội dung tại chỗ khi được. */
export interface PersistentDialog<V extends DialogView> {
  /** `true` → `update` gửi được nội dung mới vào hộp thoại đang mở (DialogApi 1.2). */
  canUpdate: boolean;
  /** Tin kế tiếp từ hộp thoại; người dùng đóng hộp thoại → `{ action: "cancel" }`. */
  next: () => Promise<DialogReplies[V]>;
  /** Gửi nội dung mới vào hộp thoại (DialogApi 1.2). `false` → bản Office cũ, cần đóng rồi mở lại. */
  update: (payload: DialogPayloads[V]) => boolean;
  close: () => void;
}

/**
 * Mở hộp thoại không tự đóng khi nhận tin (khác `openDialog`). DialogApi 1.2 (`messageChild`) cập nhật nội dung
 * tại chỗ; Excel 2016 (1.1) không có → bên gọi đóng và mở lại.
 */
export async function openPersistentDialog<V extends DialogView>(
  view: V,
  payload: DialogPayloads[V],
  size: DialogSize,
): Promise<PersistentDialog<V>> {
  const json = JSON.stringify(payload);
  const id = newId();
  const canUpdate = Office.context.requirements.isSetSupported("DialogApi", "1.2");
  const queue: DialogReplies[V][] = [];
  let waiting: ((reply: DialogReplies[V]) => void) | undefined;
  let ended = false;
  const push = (reply: DialogReplies[V]): void => {
    if (waiting) {
      const resume = waiting;
      waiting = undefined;
      resume(reply);
    } else {
      queue.push(reply);
    }
  };
  const end = (): void => {
    if (ended) return;
    ended = true;
    removePayload(id);
    push({ action: "cancel" });
  };
  const win = await openWindow(
    view,
    json,
    id,
    size,
    (message) => {
      push(parseReply(view, message));
    },
    end,
  );
  return {
    canUpdate,
    next: () =>
      new Promise((done) => {
        const queued = queue.shift();
        if (queued) done(queued);
        else waiting = done;
      }),
    update: (next) => {
      if (!canUpdate || ended || !win.current) return false;
      try {
        win.current.messageChild(JSON.stringify(next));
        return true;
      } catch {
        // Hộp thoại đang được mở lại (đường dẫn đầy đủ) → bỏ qua lần cập nhật này.
        return false;
      }
    },
    close: () => {
      if (ended) return;
      win.current?.close();
      end();
    },
  };
}
