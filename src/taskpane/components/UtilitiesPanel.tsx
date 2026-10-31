import { DefaultButton, Spinner, SpinnerSize, Stack, Text } from "@fluentui/react";
import { type JSX, useMemo } from "react";
import type { UpdateStatus } from "../../core/version/version";
import { getHostEnvironment } from "../../excel/environment";
import { APP_VERSION } from "../../version/version";
import { HelpSection } from "./utilities/HelpSection";
import { UpdateStatusMessage } from "./utilities/UpdateStatusMessage";

interface UtilitiesPanelProps {
  updateStatus: UpdateStatus | undefined;
  checking: boolean;
  onCheckUpdate: () => void;
}

/** Tab Tiện ích (F3): hướng dẫn sử dụng, phiên bản và kiểm tra cập nhật. */
export function UtilitiesPanel({ updateStatus, checking, onCheckUpdate }: UtilitiesPanelProps): JSX.Element {
  const environment = useMemo(() => getHostEnvironment(), []);

  const rows: ReadonlyArray<[string, string]> = [
    ["Phiên bản add-in", APP_VERSION],
    ["Nền tảng", environment.platform],
    ["Phiên bản Office", environment.officeVersion],
    ["ExcelApi tối đa", environment.maxExcelApi],
  ];

  return (
    <Stack tokens={{ childrenGap: 16, padding: "12px 0" }}>
      <HelpSection />

      <Stack tokens={{ childrenGap: 6 }}>
        <Text variant="large">Thông tin phiên bản</Text>
        {rows.map(([label, value]) => (
          <Text key={label}>
            <strong>{label}:</strong> {value}
          </Text>
        ))}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Kiểm tra cập nhật" onClick={onCheckUpdate} disabled={checking} />
          {checking && <Spinner size={SpinnerSize.small} />}
        </Stack>
        {updateStatus && <UpdateStatusMessage status={updateStatus} />}
      </Stack>
    </Stack>
  );
}
