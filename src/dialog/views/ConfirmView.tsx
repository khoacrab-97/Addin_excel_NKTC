import { MessageBar, MessageBarType, Text } from "@fluentui/react";
import type { JSX } from "react";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { DialogFrame } from "../components/DialogFrame";

interface Props {
  payload: DialogPayloads["confirm"];
  reply: (value: DialogReplies["confirm"]) => void;
}

/** R2.4, R2.5: xác nhận trước khi ghi đè ô đích đang có công thức hoặc dữ liệu. */
export function ConfirmView({ payload, reply }: Props): JSX.Element {
  return (
    <DialogFrame
      title="Xác nhận ghi đè"
      actions={[
        {
          text: "Ghi đè",
          primary: true,
          onClick: () => {
            reply({ action: "ok" });
          },
        },
        {
          text: "Hủy",
          onClick: () => {
            reply({ action: "cancel" });
          },
        },
      ]}
    >
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        Các ô sau đang có dữ liệu hoặc công thức và sẽ bị ghi đè bằng dữ liệu ngày {payload.day}:
      </MessageBar>
      <ul style={{ margin: 0, paddingLeft: 18 }}>
        {payload.cells.map((cell) => (
          <li key={cell.target}>
            {cell.target}
            {cell.formula ? " — có công thức" : cell.value ? " — có dữ liệu" : ""}
          </li>
        ))}
      </ul>
      <Text variant="small">
        Chỉ hỏi một lần cho bộ ô đã gán này; các lần chuyển ngày sau sẽ ghi trực tiếp. Thao tác của add-in có thể không
        hoàn tác được bằng Ctrl+Z — hãy lưu file trước nếu cần.
      </Text>
    </DialogFrame>
  );
}
