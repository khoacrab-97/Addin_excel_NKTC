import { DefaultButton, Stack, Text, TextField } from "@fluentui/react";
import { type JSX, useState } from "react";
import { formatDay } from "../../../core/dates/excelDate";

interface DayNavigatorProps {
  currentDay: number;
  position: number;
  total: number;
  disabled: boolean;
  onStep: (delta: number) => void;
  onJump: (text: string) => void;
}

/** Thanh điều hướng ngày: Đầu / Trước / Sau / Cuối và nhảy tới ngày. */
const navButtonStyles = { root: { minWidth: 36, padding: "0 6px" } };

export function DayNavigator({
  currentDay,
  position,
  total,
  disabled,
  onStep,
  onJump,
}: DayNavigatorProps): JSX.Element {
  const [jumpText, setJumpText] = useState("");
  const atStart = position <= 0;
  const atEnd = position >= total - 1;
  // Dùng ký tự thay icon: vẫn hiển thị khi máy không tải được bộ icon từ CDN của Microsoft.
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Stack horizontal verticalAlign="center" horizontalAlign="space-between" tokens={{ childrenGap: 4 }}>
        <DefaultButton
          text="«"
          styles={navButtonStyles}
          ariaLabel="Ngày đầu"
          title="Ngày đầu"
          onClick={() => {
            onStep(-Infinity);
          }}
          disabled={disabled || atStart}
        />
        <DefaultButton
          text="‹"
          styles={navButtonStyles}
          ariaLabel="Ngày trước"
          title="Ngày trước"
          onClick={() => {
            onStep(-1);
          }}
          disabled={disabled || atStart}
        />
        <Stack horizontalAlign="center">
          <Text variant="large" styles={{ root: { fontWeight: 600 } }} data-testid="current-day">
            {formatDay(currentDay)}
          </Text>
          <Text variant="small">
            Ngày {position + 1}/{total}
          </Text>
        </Stack>
        <DefaultButton
          text="›"
          styles={navButtonStyles}
          ariaLabel="Ngày sau"
          title="Ngày sau"
          onClick={() => {
            onStep(1);
          }}
          disabled={disabled || atEnd}
        />
        <DefaultButton
          text="»"
          styles={navButtonStyles}
          ariaLabel="Ngày cuối"
          title="Ngày cuối"
          onClick={() => {
            onStep(Infinity);
          }}
          disabled={disabled || atEnd}
        />
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="end">
        <TextField
          placeholder="dd/mm/yyyy"
          ariaLabel="Nhảy tới ngày"
          value={jumpText}
          onChange={(_event, value) => {
            setJumpText(value ?? "");
          }}
          styles={{ root: { flexGrow: 1 } }}
          disabled={disabled}
        />
        <DefaultButton
          text="Đi tới"
          onClick={() => {
            onJump(jumpText);
          }}
          disabled={disabled || jumpText.trim() === ""}
        />
      </Stack>
    </Stack>
  );
}
