import { MessageBar, MessageBarType, Text } from "@fluentui/react";
import { type JSX, useState } from "react";
import { type ColumnKey, type DailyLogColumns, isCompleteColumns } from "../../core/dailyLog/columns";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { ColumnDropdown } from "../components/ColumnDropdown";
import { DialogFrame } from "../components/DialogFrame";
import { RangeField } from "../components/RangeField";

const FIELDS: ReadonlyArray<{ key: ColumnKey; label: string }> = [
  { key: "content", label: "Nội dung công việc" },
  { key: "start", label: "Ngày bắt đầu" },
  { key: "end", label: "Ngày kết thúc" },
];

interface Props {
  payload: DialogPayloads["f1-columns"];
  reply: (value: DialogReplies["f1-columns"]) => void;
}

/** F1: vùng dữ liệu (tự nhận diện, nút mũi tên để chọn lại) + 3 cột của bảng tiến độ. */
export function F1ColumnsView({ payload, reply }: Props): JSX.Element {
  const [columns, setColumns] = useState<Partial<DailyLogColumns>>(payload.detected);
  const sourceOk = payload.source.error === undefined && payload.options.length > 0;
  const complete = sourceOk && isCompleteColumns(columns);

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
      <RangeField
        label="Vùng dữ liệu (gồm dòng tiêu đề)"
        value={payload.source.value}
        error={payload.source.error}
        onPick={(address) => {
          reply(address === undefined ? { action: "pick" } : { action: "pick", address });
        }}
      />
      <Text variant="small">
        Add-in tự nhận diện vùng. Muốn chọn lại: kéo chọn vùng trên sheet rồi bấm ⬆ (hoặc gõ địa chỉ rồi Enter).
      </Text>
      {sourceOk &&
        FIELDS.map((field) => (
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
      {sourceOk && !isCompleteColumns(columns) && (
        <MessageBar messageBarType={MessageBarType.warning}>Chọn đủ 3 cột khác nhau.</MessageBar>
      )}
    </DialogFrame>
  );
}
