import { PrimaryButton, DefaultButton, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";

interface Props {
  payload: DialogPayloads["pick"];
  reply: (value: DialogReplies["pick"]) => void;
}

/** Hộp thoại nhỏ của nút ⬆: người dùng chọn vùng/ô trên sheet rồi bấm OK (giống ô chọn vùng của Excel). */
export function PickView({ payload, reply }: Props): JSX.Element {
  return (
    <Stack styles={{ root: { padding: 12 } }} tokens={{ childrenGap: 10 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        Chọn {payload.label}
      </Text>
      <Text>
        {payload.kind === "cell"
          ? "Bấm chọn ô (hoặc ô gộp) trên sheet, rồi bấm OK."
          : "Kéo chọn vùng trên sheet (hoặc bấm 1 ô trong bảng), rồi bấm OK."}
      </Text>
      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text="OK"
          onClick={() => {
            reply({ action: "ok" });
          }}
        />
        <DefaultButton
          text="Hủy"
          onClick={() => {
            reply({ action: "cancel" });
          }}
        />
      </Stack>
    </Stack>
  );
}
