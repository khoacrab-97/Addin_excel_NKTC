import { Stack, Text } from "@fluentui/react";
import { type JSX, useMemo } from "react";
import { getHostEnvironment } from "../../excel/environment";
import { APP_VERSION } from "../../version/version";

/** Tab Tiện ích: hiện tại hiển thị phiên bản và môi trường (phục vụ kiểm tra P0.6). */
export function UtilitiesPanel(): JSX.Element {
  const environment = useMemo(() => getHostEnvironment(), []);

  const rows: ReadonlyArray<[string, string]> = [
    ["Phiên bản add-in", APP_VERSION],
    ["Nền tảng", environment.platform],
    ["Phiên bản Office", environment.officeVersion],
    ["ExcelApi tối đa", environment.maxExcelApi],
  ];

  return (
    <Stack tokens={{ childrenGap: 6, padding: "12px 0" }}>
      <Text variant="large">Thông tin phiên bản</Text>
      {rows.map(([label, value]) => (
        <Text key={label}>
          <strong>{label}:</strong> {value}
        </Text>
      ))}
    </Stack>
  );
}
