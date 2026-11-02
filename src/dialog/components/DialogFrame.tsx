import { DefaultButton, PrimaryButton, Stack, Text } from "@fluentui/react";
import type { JSX, ReactNode } from "react";

export interface FrameAction {
  text: string;
  onClick: () => void;
  primary?: boolean;
  disabled?: boolean;
}

interface DialogFrameProps {
  title: string;
  children: ReactNode;
  actions: FrameAction[];
}

/** Khung chung của hộp thoại: tiêu đề, nội dung cuộn được, các nút ở cuối. */
export function DialogFrame({ title, children, actions }: DialogFrameProps): JSX.Element {
  return (
    <Stack styles={{ root: { height: "100vh", boxSizing: "border-box", padding: 16 } }} tokens={{ childrenGap: 12 }}>
      <Text variant="xLarge">{title}</Text>
      <Stack grow styles={{ root: { overflowY: "auto", minHeight: 0 } }} tokens={{ childrenGap: 10 }}>
        {children}
      </Stack>
      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
        {actions.map((action) =>
          action.primary ? (
            <PrimaryButton
              key={action.text}
              text={action.text}
              onClick={action.onClick}
              disabled={action.disabled === true}
            />
          ) : (
            <DefaultButton
              key={action.text}
              text={action.text}
              onClick={action.onClick}
              disabled={action.disabled === true}
            />
          ),
        )}
      </Stack>
    </Stack>
  );
}
