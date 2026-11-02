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

/** URL thư mục chứa function file (dialog.html nằm cùng thư mục, cùng domain — yêu cầu của Dialog API). */
function baseUrl(): string {
  return window.location.href.replace(/[?#].*$/, "").replace(/[^/]*$/, "");
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
  const id = `${String(Date.now())}${Math.random().toString(36).slice(2, 8)}`;
  storePayload(id, json);

  let url = `${baseUrl()}dialog.html?view=${view}&id=${id}`;
  const hash = encodeURIComponent(json);
  if (url.length + hash.length + 1 <= MAX_URL_LENGTH) url += `#${hash}`;

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
          const message = "message" in args ? args.message : "";
          dialog.close();
          finish(parseReply(view, message));
        });
        // Người dùng đóng hộp thoại (12006) hoặc trang hộp thoại lỗi.
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          finish({ action: "cancel" });
        });
      },
    );
  });
}
