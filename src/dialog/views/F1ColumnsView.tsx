import { MessageBar, MessageBarType, Text } from "@fluentui/react";
import { type JSX, useState } from "react";
import { type ColumnKey, type DailyLogColumns, isCompleteColumns } from "../../core/dailyLog/columns";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { ColumnDropdown } from "../components/ColumnDropdown";
import { DialogFrame } from "../components/DialogFrame";

const FIELDS: ReadonlyArray<{ key: ColumnKey; label: string }> = [
  { key: "content", label: "Nội dung công việc" },
  { key: "start", label: "Ngày bắt đầu" },
  { key: "end", label: "Ngày kết thúc" },
];

interface Props {
  payload: DialogPayloads["f1-columns"];
  reply: (value: DialogReplies["f1-columns"]) => void;
}

/** F1: xác nhận 3 cột của bảng tiến độ trước khi xuất. */
export function F1ColumnsView({ payload, reply }: Props): JSX.Element {
  const [columns, setColumns] = useState<Partial<DailyLogColumns>>(payload.detected);
  const complete = isCompleteColumns(columns);

  return (
    <DialogFrame
      title="Xuất nhật ký theo ngày"
      actions={[
        {
          text: "Xuất nhật ký",
          primary: true,
          disabled: !complete,
          onClick: () => {
            if (isCompleteColumns(columns)) reply({ action: "ok", columns });
          },
        },
        {
          text: "Hủy",
          onClick: () => {
            reply({ action: "cancel" });
          },
        },
      ]}
    >
      <Text>
        <strong>Vùng dữ liệu:</strong> {payload.source}
      </Text>
      <Text variant="small">Kiểm tra 3 cột được nhận diện, chỉnh lại nếu cần. Kết quả ghi vào sheet mới.</Text>
      {FIELDS.map((field) => (
        <ColumnDropdown
          key={field.key}
          label={field.label}
          options={payload.options}
          selected={columns[field.key]}
          onChange={(column) => {
            setColumns((previous) => ({ ...previous, [field.key]: column }));
          }}
        />
      ))}
      {!complete && <MessageBar messageBarType={MessageBarType.warning}>Chọn đủ 3 cột khác nhau.</MessageBar>}
    </DialogFrame>
  );
}
