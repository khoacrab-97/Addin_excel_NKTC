import { Text, TextField } from "@fluentui/react";
import { type JSX, useState } from "react";
import { parseDateText } from "../../core/dates/excelDate";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { DialogFrame } from "../components/DialogFrame";

interface Props {
  payload: DialogPayloads["goto"];
  reply: (value: DialogReplies["goto"]) => void;
}

/** F2: nhập ngày cần hiển thị trên biểu mẫu. */
export function GotoView({ payload, reply }: Props): JSX.Element {
  const [text, setText] = useState(payload.current ?? "");
  const valid = parseDateText(text) !== undefined;
  const submit = (): void => {
    if (valid) reply({ action: "goto", text: text.trim() });
  };

  return (
    <DialogFrame
      title="Đến ngày"
      actions={[
        { text: "Đi tới", primary: true, disabled: !valid, onClick: submit },
        {
          text: "Hủy",
          onClick: () => {
            reply({ action: "cancel" });
          },
        },
      ]}
    >
      {payload.first && payload.last && (
        <Text>
          Vùng nguồn có dữ liệu từ {payload.first} đến {payload.last}.
        </Text>
      )}
      <TextField
        label="Ngày (dd/mm/yyyy)"
        placeholder="dd/mm/yyyy"
        value={text}
        autoFocus
        onChange={(_event, value) => {
          setText(value ?? "");
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter") submit();
        }}
        errorMessage={text.trim() !== "" && !valid ? "Ngày không hợp lệ." : ""}
      />
    </DialogFrame>
  );
}
