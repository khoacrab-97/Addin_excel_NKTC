import { DefaultButton, Spinner, SpinnerSize, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { APP_VERSION } from "../../version/version";
import { DialogFrame } from "../components/DialogFrame";
import { UpdateStatusMessage } from "../components/UpdateStatusMessage";
import { useUpdateCheck } from "../components/useUpdateCheck";

interface Props {
  payload: DialogPayloads["about"];
  reply: (value: DialogReplies["about"]) => void;
}

/** F3: thông tin phiên bản và kiểm tra cập nhật (P3.2, P3.3). */
export function AboutView({ payload, reply }: Props): JSX.Element {
  const update = useUpdateCheck();
  const rows: ReadonlyArray<[string, string]> = [
    ["Phiên bản add-in", APP_VERSION],
    ["Nền tảng", payload.platform],
    ["Phiên bản Office", payload.officeVersion],
    ["ExcelApi tối đa", payload.maxExcelApi],
  ];

  return (
    <DialogFrame
      title="Phiên bản và cập nhật"
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
      {rows.map(([label, value]) => (
        <Text key={label}>
          <strong>{label}:</strong> {value}
        </Text>
      ))}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Kiểm tra lại" onClick={update.recheck} disabled={update.checking} />
        {(update.checking || !update.status) && <Spinner size={SpinnerSize.small} />}
      </Stack>
      {update.status && <UpdateStatusMessage status={update.status} />}
    </DialogFrame>
  );
}
