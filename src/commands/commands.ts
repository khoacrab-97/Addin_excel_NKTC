import { getHostEnvironment } from "../excel/environment";
import { exportDailyLogCommand } from "./dailyLogCommand";
import { DIALOG_SIZE, openDialog } from "./dialogHost";
import { showError } from "./messages";
import {
  assignCellCommand,
  goToDayCommand,
  mappingListCommand,
  setSourceCommand,
  showDayCommand,
} from "./templateCommands";

/** Một lệnh của nút trên Ribbon (Action ExecuteFunction trong manifest). */
type CommandHandler = () => Promise<void>;

/** Tên hàm phải trùng `<FunctionName>` trong manifest.xml. */
export const COMMANDS: Readonly<Record<string, CommandHandler>> = {
  nktcExportDaily: exportDailyLogCommand,
  nktcSetSource: setSourceCommand,
  nktcAssignCell: assignCellCommand,
  nktcMappingList: mappingListCommand,
  nktcDayFirst: () => showDayCommand({ kind: "first" }),
  nktcDayPrev: () => showDayCommand({ kind: "prev" }),
  nktcDayNext: () => showDayCommand({ kind: "next" }),
  nktcDayLast: () => showDayCommand({ kind: "last" }),
  nktcGoToDay: goToDayCommand,
  nktcHelp: async () => {
    await openDialog("help", {}, DIALOG_SIZE.large);
  },
  nktcAbout: async () => {
    await openDialog("about", getHostEnvironment(), DIALOG_SIZE.medium);
  },
};

/** Chỉ chạy một lệnh tại một thời điểm (bấm liên tiếp khi đang mở hộp thoại → bỏ qua). */
let running = false;

function wrap(handler: CommandHandler): (event: Office.AddinCommands.Event) => void {
  return (event) => {
    if (running) {
      event.completed();
      return;
    }
    running = true;
    handler()
      .catch((error: unknown) => showError(error))
      .finally(() => {
        running = false;
        event.completed();
      });
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
