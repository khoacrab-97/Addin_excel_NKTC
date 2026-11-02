import { MessageBar, MessageBarType, Stack } from "@fluentui/react";
import type { JSX } from "react";
import type { DialogPayloads, DialogReplies, MessageTone } from "../../core/dialog/protocol";
import { DialogFrame } from "../components/DialogFrame";

const TONES: Readonly<Record<MessageTone, MessageBarType>> = {
  info: MessageBarType.info,
  success: MessageBarType.success,
  warning: MessageBarType.warning,
  error: MessageBarType.error,
};

const cellStyle = { border: "1px solid #c8c8c8", padding: "4px 6px", verticalAlign: "top" } as const;

interface Props {
  payload: DialogPayloads["message"];
  reply: (value: DialogReplies["message"]) => void;
}

/** Thông báo kết quả / lỗi của một lệnh. */
export function MessageView({ payload, reply }: Props): JSX.Element {
  const [first, ...rest] = payload.lines;
  return (
    <DialogFrame
      title={payload.title}
      actions={[
        {
          text: "Đóng",
          primary: true,
          onClick: () => {
            reply({ action: "ok" });
          },
        },
      ]}
    >
      <MessageBar messageBarType={TONES[payload.tone]} isMultiline>
        {first}
      </MessageBar>
      {rest.length > 0 && (
        <Stack tokens={{ childrenGap: 4 }}>
          {rest.map((line, index) => (
            <div key={index}>{line}</div>
          ))}
        </Stack>
      )}
      {payload.table && (
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
          <thead>
            <tr>
              {payload.table.headers.map((header) => (
                <th key={header} style={cellStyle}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {payload.table.rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex} style={{ ...cellStyle, whiteSpace: "pre-wrap" }}>
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </DialogFrame>
  );
}
