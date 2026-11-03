import { DefaultButton, Stack, TextField } from "@fluentui/react";
import { type JSX, useState } from "react";

interface RangeFieldProps {
  label: string;
  /** Giá trị đang áp dụng (do add-in đọc từ sheet). */
  value: string;
  error?: string | undefined;
  placeholder?: string;
  /** Thu gọn để chọn trên sheet (`address` không có) hoặc áp dụng địa chỉ gõ tay. */
  onPick: (address?: string) => void;
  ariaLabel?: string;
  /** Có thì giá trị gõ tay được nhận ngay (không cần Enter), ví dụ ô đích trong danh sách gán. */
  onTextChange?: (text: string) => void;
}

/**
 * Ô chọn vùng giống Excel: hiển thị vùng đang dùng (tự nhận diện), nút mũi tên ở cuối thu gọn hộp thoại thành
 * thanh chọn để chọn trên sheet. Có thể gõ địa chỉ rồi Enter.
 */
export function RangeField({
  label,
  value,
  error,
  placeholder,
  onPick,
  ariaLabel,
  onTextChange,
}: RangeFieldProps): JSX.Element {
  const [text, setText] = useState(value);
  const edited = onTextChange === undefined && text.trim() !== value;
  return (
    <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 4 }}>
      <Stack.Item grow>
        <TextField
          label={label}
          value={text}
          placeholder={placeholder ?? "Sheet!A1:D20"}
          ariaLabel={ariaLabel ?? label}
          onChange={(_event, next) => {
            setText(next ?? "");
            onTextChange?.(next ?? "");
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter" && text.trim() !== "") onPick(text.trim());
          }}
          errorMessage={error ?? (edited ? "Nhấn Enter để áp dụng địa chỉ vừa gõ." : "")}
        />
      </Stack.Item>
      <DefaultButton
        text="⬆"
        title="Thu gọn hộp thoại để chọn trên sheet"
        ariaLabel={`Chọn ${ariaLabel ?? label} trên sheet`}
        styles={{ root: { minWidth: 36, padding: 0, marginBottom: error || edited ? 22 : 0 } }}
        onClick={() => {
          onPick();
        }}
      />
    </Stack>
  );
}
