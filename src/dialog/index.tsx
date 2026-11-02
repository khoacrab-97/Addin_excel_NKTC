import { initializeIcons } from "@fluentui/react";
import type { JSX } from "react";
import ReactDOM from "react-dom";
import { type LoadedDialog, loadDialog } from "./payload";
import { AboutView } from "./views/AboutView";
import { AssignView } from "./views/AssignView";
import { ConfirmView } from "./views/ConfirmView";
import { F1ColumnsView } from "./views/F1ColumnsView";
import { GotoView } from "./views/GotoView";
import { HelpView } from "./views/HelpView";
import { MessageView } from "./views/MessageView";
import { SettingsView } from "./views/SettingsView";

/** Gửi kết quả về nút lệnh trên Ribbon; function file đóng hộp thoại khi nhận được. */
function reply(value: object): void {
  Office.context.ui.messageParent(JSON.stringify(value));
}

function DialogApp({ dialog }: { dialog: LoadedDialog }): JSX.Element {
  switch (dialog.view) {
    case "f1-columns":
      return <F1ColumnsView payload={dialog.payload} reply={reply} />;
    case "message":
      return <MessageView payload={dialog.payload} reply={reply} />;
    case "assign":
      return <AssignView payload={dialog.payload} reply={reply} />;
    case "settings":
      return <SettingsView payload={dialog.payload} reply={reply} />;
    case "goto":
      return <GotoView payload={dialog.payload} reply={reply} />;
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
