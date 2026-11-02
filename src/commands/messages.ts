import type { MessagePayload, MessageTone } from "../core/dialog/protocol";
import { toUserMessage } from "../excel/errors";
import { DIALOG_SIZE, openDialog } from "./dialogHost";

/** Hiển thị thông báo trong hộp thoại nhỏ và chờ người dùng bấm Đóng. */
export async function showMessage(
  title: string,
  tone: MessageTone,
  lines: string[],
  table?: MessagePayload["table"],
): Promise<void> {
  const payload: MessagePayload = table ? { title, tone, lines, table } : { title, tone, lines };
  await openDialog("message", payload, table ? DIALOG_SIZE.large : DIALOG_SIZE.small);
}

/** Báo lỗi của một lệnh. Nếu không mở được hộp thoại (ví dụ đang có hộp thoại khác) thì bỏ qua. */
export async function showError(error: unknown, title = "Không thực hiện được"): Promise<void> {
  try {
    await showMessage(title, "error", [toUserMessage(error)]);
  } catch {
    // Không còn cách hiển thị nào khác từ nút lệnh trên Ribbon.
  }
}
