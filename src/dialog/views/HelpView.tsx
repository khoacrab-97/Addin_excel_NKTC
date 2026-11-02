import type { JSX } from "react";
import type { DialogReplies } from "../../core/dialog/protocol";
import { DialogFrame } from "../components/DialogFrame";
import { HelpSection } from "../components/HelpSection";

interface Props {
  reply: (value: DialogReplies["help"]) => void;
}

/** F3: hướng dẫn sử dụng và file mẫu (P3.1). */
export function HelpView({ reply }: Props): JSX.Element {
  return (
    <DialogFrame
      title="Hướng dẫn sử dụng"
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
      <HelpSection />
    </DialogFrame>
  );
}
