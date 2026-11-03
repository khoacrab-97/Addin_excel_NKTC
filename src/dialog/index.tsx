import { initializeIcons } from "@fluentui/react";
import type { JSX } from "react";
import ReactDOM from "react-dom";
import { NO_PAYLOAD_MESSAGE } from "../core/dialog/protocol";
import { type LoadedDialog, isShortUrl, loadDialog } from "./payload";
import { AboutView } from "./views/AboutView";
import { ConfirmView } from "./views/ConfirmView";
import { CreateTableView } from "./views/CreateTableView";
import { ExportFormView } from "./views/ExportFormView";
import { F1ColumnsView } from "./views/F1ColumnsView";
import { HelpView } from "./views/HelpView";
import { MessageView } from "./views/MessageView";
import { PickView } from "./views/PickView";

/** Gửi kết quả về nút lệnh trên Ribbon; function file đóng hộp thoại khi nhận được. */
function reply(value: object): void {
  Office.context.ui.messageParent(JSON.stringify(value));
}

function DialogApp({ dialog }: { dialog: LoadedDialog }): JSX.Element {
  switch (dialog.view) {
    case "create-table":
      return <CreateTableView payload={dialog.payload} reply={reply} />;
    case "f1-columns":
      return <F1ColumnsView payload={dialog.payload} reply={reply} />;
    case "message":
      return <MessageView payload={dialog.payload} reply={reply} />;
    case "export-form":
      return <ExportFormView payload={dialog.payload} reply={reply} />;
    case "pick":
      return <PickView payload={dialog.payload} reply={reply} />;
    case "confirm":
      return <ConfirmView payload={dialog.payload} reply={reply} />;
    case "help":
      return <HelpView reply={reply} />;
    case "about":
      return <AboutView payload={dialog.payload} reply={reply} />;
  }
}

Office.onReady()
  .then(() => {
    initializeIcons();
    const root = document.getElementById("root");
    const dialog = loadDialog();
    if (!dialog) {
      if (isShortUrl()) {
        // Bộ nhớ không dùng chung với function file → xin mở lại bằng đường dẫn đầy đủ.
        if (root) root.textContent = "Đang mở hộp thoại…";
        Office.context.ui.messageParent(NO_PAYLOAD_MESSAGE);
        return;
      }
      if (root) root.textContent = "Không đọc được dữ liệu của hộp thoại. Vui lòng đóng hộp thoại và thử lại.";
      return;
    }
    // React 17: dùng ReactDOM.render để tương thích Internet Explorer 11 (Excel 2016/2019).
    ReactDOM.render(<DialogApp dialog={dialog} />, root);
  })
  .catch((error: unknown) => {
    const root = document.getElementById("root");
    if (root) root.textContent = `Không mở được hộp thoại: ${String(error)}`;
  });
