import { DefaultButton, MessageBar, MessageBarType, Stack, Text, TextField } from "@fluentui/react";
import { type JSX, useEffect, useState } from "react";
import { parseDateText } from "../../core/dates/excelDate";
import { type DialogPayloads, type DialogReplies, type MessageTone, parsePayload } from "../../core/dialog/protocol";

const TONES: Readonly<Record<MessageTone, MessageBarType>> = {
  info: MessageBarType.info,
  success: MessageBarType.success,
  warning: MessageBarType.warning,
  error: MessageBarType.error,
};

type NavKind = "first" | "prev" | "next" | "last";
const NAV: ReadonlyArray<{ kind: NavKind; text: string }> = [
  { kind: "first", text: "|◀ Đầu" },
  { kind: "prev", text: "◀ Trước" },
  { kind: "next", text: "Sau ▶" },
  { kind: "last", text: "Cuối ▶|" },
];

interface Props {
  payload: DialogPayloads["day-panel"];
  reply: (value: DialogReplies["day-panel"]) => void;
}

/**
 * Bảng chuyển ngày nổi (R2.16): ngày đang xem, ô gõ ngày + Enter, nút Đầu / Trước / Sau / Cuối.
 * DialogApi 1.2: add-in gửi nội dung mới vào bảng sau mỗi lần chuyển (bảng không đóng).
 */
export function DayPanelView({ payload: initial, reply }: Props): JSX.Element {
  const [payload, setPayload] = useState(initial);
  const [busy, setBusy] = useState(false);
  const [text, setText] = useState(initial.current?.value ?? "");

  useEffect(() => {
    const ui = Office.context.ui as Partial<Office.UI>;
    if (!ui.addHandlerAsync) return;
    ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (args) => {
      const next = parsePayload("day-panel", args.message);
      if (!next) return;
      setPayload(next);
      setText(next.current?.value ?? "");
      setBusy(false);
    });
  }, []);

  const send = (value: DialogReplies["day-panel"]): void => {
    setBusy(true);
    reply(value);
  };
  const valid = parseDateText(text) !== undefined;
  const locked = busy || payload.disabled === true;

  return (
    <Stack styles={{ root: { padding: 12, height: "100vh", boxSizing: "border-box" } }} tokens={{ childrenGap: 10 }}>
      <Stack tokens={{ childrenGap: 2 }}>
        <Text variant="small">Đang xem</Text>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }} data-testid="current-day">
          {payload.current?.label ?? "(chưa hiển thị ngày nào)"}
        </Text>
        {payload.range && (
          <Text variant="small">
            {payload.current?.position !== undefined
              ? `Ngày ${payload.current.position} / ${payload.range.count} · `
              : ""}
            từ {payload.range.first} đến {payload.range.last}
          </Text>
        )}
      </Stack>
      <TextField
        label="Đến ngày (gõ rồi Enter)"
        value={text}
        placeholder="dd/mm/yyyy"
        disabled={locked}
        onChange={(_event, value) => {
          setText(value ?? "");
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter" && valid) send({ action: "goto", text: text.trim() });
        }}
        errorMessage={text.trim() !== "" && !valid ? "Ngày không hợp lệ." : ""}
      />
      <Stack horizontal tokens={{ childrenGap: 6 }}>
        {NAV.map((item) => (
          <DefaultButton
            key={item.kind}
            text={item.text}
            disabled={locked}
            styles={{ root: { minWidth: 0, flex: 1, padding: "0 4px" } }}
            onClick={() => {
              send({ action: "nav", kind: item.kind });
            }}
          />
        ))}
      </Stack>
      {payload.message && (
        <MessageBar messageBarType={TONES[payload.message.tone]} isMultiline>
          {payload.message.text}
        </MessageBar>
      )}
      <Stack.Item grow>
        <span />
      </Stack.Item>
      <Stack horizontal horizontalAlign="end">
        <DefaultButton
          text="Đóng"
          onClick={() => {
            reply({ action: "cancel" });
          }}
        />
      </Stack>
    </Stack>
  );
}
