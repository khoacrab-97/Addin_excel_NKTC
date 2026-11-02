import { MessageBar, MessageBarType, Text } from "@fluentui/react";
import { type JSX, useState } from "react";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { ColumnDropdown } from "../components/ColumnDropdown";
import { DialogFrame } from "../components/DialogFrame";
import type { FrameAction } from "../components/DialogFrame";

interface Props {
  payload: DialogPayloads["assign"];
  reply: (value: DialogReplies["assign"]) => void;
}

/** F2: gán một cột của vùng nguồn cho ô đang chọn trên biểu mẫu. */
export function AssignView({ payload, reply }: Props): JSX.Element {
  const [column, setColumn] = useState<number | undefined>(payload.current);
  const actions: FrameAction[] = [
    {
      text: "Gán",
      primary: true,
      disabled: column === undefined,
      onClick: () => {
        if (column !== undefined) reply({ action: "assign", column });
      },
    },
  ];
  if (payload.current !== undefined)
    actions.push({
      text: "Bỏ gán ô này",
      onClick: () => {
        reply({ action: "unassign" });
      },
    });
  actions.push({
    text: "Hủy",
    onClick: () => {
      reply({ action: "cancel" });
    },
  });

  return (
    <DialogFrame title="Gán cột cho ô đang chọn" actions={actions}>
      <Text>
        <strong>Ô trên biểu mẫu:</strong> {payload.target}
      </Text>
      <Text>
        <strong>Vùng nguồn:</strong> {payload.source}
      </Text>
      <ColumnDropdown label="Cột ghi vào ô này" options={payload.options} selected={column} onChange={setColumn} />
      {column === payload.dateColumn && (
        <MessageBar messageBarType={MessageBarType.info}>Đây là cột Ngày — ô này sẽ hiển thị ngày đang xem.</MessageBar>
      )}
      {payload.current !== undefined && (
        <Text variant="small">Ô này đang được gán; chọn cột khác để thay, hoặc bỏ gán.</Text>
      )}
    </DialogFrame>
  );
}
