import { getHostEnvironment } from "../excel/environment";
import { createTableCommand } from "./createTableCommand";
import { exportDailyLogCommand } from "./dailyLogCommand";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showError } from "./messages";
import { exportFormCommand, insertTemplateCommand } from "./formCommands";
import { dayPanelCommand, interruptDayPanel, showDayCommand } from "./templateCommands";

/** Một lệnh của nút trên Ribbon (Action ExecuteFunction trong manifest). */
type CommandHandler = () => Promise<void>;

/** Tên hàm phải trùng `<FunctionName>` trong manifest.xml. */
export const COMMANDS: Readonly<Record<string, CommandHandler>> = {
  nktcCreateTable: createTableCommand,
  nktcExportDaily: exportDailyLogCommand,
  nktcInsertTemplate: insertTemplateCommand,
  nktcExportForm: exportFormCommand,
  nktcDayPanel: dayPanelCommand,
  nktcDayFirst: async () => {
    await showDayCommand({ kind: "first" });
  },
  nktcDayPrev: async () => {
    await showDayCommand({ kind: "prev" });
  },
  nktcDayNext: async () => {
    await showDayCommand({ kind: "next" });
  },
  nktcDayLast: async () => {
    await showDayCommand({ kind: "last" });
  },
  nktcHelp: async () => {
    await openDialog("help", {}, DIALOG_SIZE.large);
  },
  nktcAbout: async () => {
    await openDialog("about", getHostEnvironment(), DIALOG_SIZE.medium);
  },
};

/** Lệnh đang chạy. Chỉ chạy một lệnh tại một thời điểm (bấm khi đang mở hộp thoại → bỏ qua). */
let current: Promise<void> | undefined;

function run(handler: CommandHandler, event: Office.AddinCommands.Event): void {
  const task = handler()
    .catch((error: unknown) => showError(error))
    .finally(() => {
      if (current === task) current = undefined;
      event.completed();
    });
  current = task;
}

function wrap(handler: CommandHandler): (event: Office.AddinCommands.Event) => void {
  return (event) => {
    const active = current;
    if (!active) {
      run(handler, event);
      return;
    }
    // Bảng chuyển ngày để mở lâu: bấm nút khác trên Ribbon → đóng bảng rồi chạy lệnh mới.
    if (interruptDayPanel()) {
      void active.then(() => {
        if (current) event.completed();
        else run(handler, event);
      });
      return;
    }
    event.completed();
  };
}

const registry = window as unknown as Record<string, unknown>;
const office = Office as Partial<typeof Office>;
for (const name of Object.keys(COMMANDS)) {
  const handler = COMMANDS[name];
  if (!handler) continue;
  const action = wrap(handler);
  // Excel 2016: function file gọi hàm toàn cục theo tên trong manifest.
  registry[name] = action;
}

void Office.onReady().then(() => {
  // Bản Office mới: đăng ký thêm qua Office.actions (không có trên bản cũ).
  if (!office.actions) return;
  for (const name of Object.keys(COMMANDS)) {
    const action = registry[name];
    if (typeof action === "function") office.actions.associate(name, action as (event: unknown) => void);
  }
});
