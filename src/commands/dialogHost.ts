import { type DialogPayloads, type DialogReplies, type DialogView, parseReply } from "../core/dialog/protocol";
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
  /** Bảng chuyển ngày nổi. */
  panel: { width: 28, height: 42 },
  /** Hộp thoại nhỏ "chọn vùng/ô trên sheet rồi bấm OK". */
  pick: { width: 24, height: 28 },
} as const satisfies Record<string, DialogSize>;

/** Tiền tố khóa localStorage chứa dữ liệu đầu vào của hộp thoại. */
export const PAYLOAD_STORAGE_PREFIX = "nktc.dialog.";
/** URL hộp thoại trên Internet Explorer 11 (Excel 2016) tối đa ~2.000 ký tự. */
const MAX_URL_LENGTH = 1_900;

function storePayload(id: string, json: string): void {
  try {
    window.localStorage.setItem(PAYLOAD_STORAGE_PREFIX + id, json);
  } catch {
    // localStorage bị chặn: hộp thoại vẫn đọc được dữ liệu từ URL nếu đủ ngắn.
  }
}

function removePayload(id: string): void {
  try {
    window.localStorage.removeItem(PAYLOAD_STORAGE_PREFIX + id);
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

function newId(): string {
  return `${String(Date.now())}${Math.random().toString(36).slice(2, 8)}`;
}

function dialogUrl(view: DialogView, json: string, id: string): string {
  let url = `${baseUrl()}dialog.html?view=${view}&id=${id}`;
  const hash = encodeURIComponent(json);
  if (url.length + hash.length + 1 <= MAX_URL_LENGTH) url += `#${hash}`;
  return url;
}

/**
 * Mở hộp thoại (Office Dialog API 1.1) và chờ kết quả. Người dùng đóng hộp thoại bằng nút X → `{ action: "cancel" }`.
 * Dữ liệu đầu vào được gửi qua URL (nếu đủ ngắn) và localStorage (cùng domain với hộp thoại).
 */
export function openDialog<V extends DialogView>(
  view: V,
  payload: DialogPayloads[V],
  size: DialogSize = DIALOG_SIZE.medium,
): Promise<DialogReplies[V]> {
  const json = JSON.stringify(payload);
  const id = newId();
  storePayload(id, json);
  const url = dialogUrl(view, json, id);

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { width: size.width, height: size.height, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          removePayload(id);
          reject(new UserFacingError(dialogErrorMessage(result.error.code)));
          return;
        }
        const dialog = result.value;
        let settled = false;
        const finish = (reply: DialogReplies[V]): void => {
          if (settled) return;
          settled = true;
          removePayload(id);
          resolve(reply);
        };
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          if (!("message" in args)) return;
          // Chỉ nhận kết quả từ trang hộp thoại của add-in. `origin` không có trên bản Office cũ
          // (DialogOrigin 1.1) — khi đó Office chỉ chuyển tin từ trang cùng domain đã mở.
          if (args.origin !== undefined && args.origin !== addinOrigin()) return;
          dialog.close();
          finish(parseReply(view, args.message));
        });
        // Người dùng đóng hộp thoại (12006) hoặc trang hộp thoại lỗi.
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          finish({ action: "cancel" });
        });
      },
    );
  });
}

/** Hộp thoại mở liên tục (bảng chuyển ngày): nhận nhiều tin, cập nhật nội dung tại chỗ khi được. */
export interface PersistentDialog<V extends DialogView> {
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
export function openPersistentDialog<V extends DialogView>(
  view: V,
  payload: DialogPayloads[V],
  size: DialogSize,
): Promise<PersistentDialog<V>> {
  const json = JSON.stringify(payload);
  const id = newId();
  storePayload(id, json);
  const url = dialogUrl(view, json, id);
  const canUpdate = Office.context.requirements.isSetSupported("DialogApi", "1.2");

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { width: size.width, height: size.height, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          removePayload(id);
          reject(new UserFacingError(dialogErrorMessage(result.error.code)));
          return;
        }
        const dialog = result.value;
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
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          if (!("message" in args)) return;
          if (args.origin !== undefined && args.origin !== addinOrigin()) return;
          push(parseReply(view, args.message));
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, end);
        resolve({
          next: () =>
            new Promise((done) => {
              const queued = queue.shift();
              if (queued) done(queued);
              else waiting = done;
            }),
          update: (next) => {
            if (!canUpdate || ended) return false;
            dialog.messageChild(JSON.stringify(next));
            return true;
          },
          close: () => {
            if (ended) return;
            dialog.close();
            end();
          },
        });
      },
    );
  });
}
