import { DefaultButton, Stack, Text, TextField } from "@fluentui/react";
import { type JSX, useEffect, useState } from "react";
import { type DialogPayloads, type DialogReplies, parsePayload } from "../../core/dialog/protocol";

interface Props {
  payload: DialogPayloads["pick"];
  reply: (value: DialogReplies["pick"]) => void;
}

/**
 * Thanh chọn (giống hộp thoại Excel khi thu gọn, ví dụ New Formatting Rule): tiêu đề, ô nhập, nút mũi tên ở cuối.
 * Người dùng chọn trên sheet → add-in gửi địa chỉ mới vào ô nhập (DialogApi 1.2). Bấm mũi tên hoặc Enter → xác
 * nhận, quay lại hộp thoại ban đầu. Đóng thanh (X) → hủy, giữ giá trị cũ.
 */
export function PickView({ payload, reply }: Props): JSX.Element {
  const [text, setText] = useState(payload.value);
  const [typed, setTyped] = useState(false);
  const [sent, setSent] = useState(false);

  useEffect(() => {
    const ui = Office.context.ui as Partial<Office.UI>;
    if (!ui.addHandlerAsync) return;
    ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (args) => {
      const next = parsePayload("pick", args.message);
      if (!next) return;
      setText(next.value);
      setTyped(false);
    });
  }, []);

  const confirm = (): void => {
    if (sent) return;
    setSent(true);
    reply({ action: "ok", text: text.trim(), typed: typed && text.trim() !== "" });
  };

  return (
    <Stack styles={{ root: { padding: "8px 10px" } }} tokens={{ childrenGap: 6 }}>
      <Text styles={{ root: { fontWeight: 600 } }}>{payload.title}</Text>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
        <Stack.Item grow>
          <TextField
            value={text}
            autoFocus
            placeholder={payload.kind === "cell" ? "Sheet!E2" : "Sheet!A1:D20"}
            ariaLabel="Địa chỉ đang chọn"
            onChange={(_event, next) => {
              setText(next ?? "");
              setTyped(true);
            }}
            onKeyDown={(event) => {
              if (event.key === "Enter") confirm();
            }}
          />
        </Stack.Item>
        <DefaultButton
          text="⬇"
          title="Xác nhận và quay lại hộp thoại"
          ariaLabel="Xác nhận vùng chọn"
          disabled={sent}
          styles={{ root: { minWidth: 36, padding: 0 } }}
          onClick={confirm}
        />
      </Stack>
      <Text variant="small">
        {payload.kind === "cell"
          ? "Bấm chọn ô trên sheet, rồi bấm ⬇ (hoặc Enter) để quay lại."
          : "Kéo chọn vùng trên sheet, rồi bấm ⬇ (hoặc Enter) để quay lại."}
      </Text>
    </Stack>
  );
}
