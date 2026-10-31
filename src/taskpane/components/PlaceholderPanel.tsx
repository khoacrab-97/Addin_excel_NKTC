import type { JSX } from "react";
import { MessageBar, MessageBarType, Stack, Text } from "@fluentui/react";

interface PlaceholderPanelProps {
  title: string;
  phase: string;
}

/** Khung tạm cho chức năng chưa triển khai. */
export function PlaceholderPanel({ title, phase }: PlaceholderPanelProps): JSX.Element {
  return (
    <Stack tokens={{ childrenGap: 8, padding: "12px 0" }}>
      <Text variant="large">{title}</Text>
      <MessageBar messageBarType={MessageBarType.info}>Chức năng sẽ được triển khai ở {phase}.</MessageBar>
    </Stack>
  );
}
