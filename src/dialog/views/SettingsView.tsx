import { Checkbox, Text } from "@fluentui/react";
import { type JSX, useState } from "react";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { ColumnDropdown } from "../components/ColumnDropdown";
import { DialogFrame } from "../components/DialogFrame";

const cellStyle = { border: "1px solid #c8c8c8", padding: "4px 6px", verticalAlign: "middle" } as const;

interface Props {
  payload: DialogPayloads["settings"];
  reply: (value: DialogReplies["settings"]) => void;
}

/** F2: danh sách ô đã gán — đổi cột Ngày, bỏ gán ô. */
export function SettingsView({ payload, reply }: Props): JSX.Element {
  const [dateColumn, setDateColumn] = useState(payload.dateColumn);
  const [remove, setRemove] = useState<string[]>([]);
  const labelOf = (column: number): string =>
    payload.options.find((option) => option.index === column)?.label ?? `Cột ${column + 1}`;

  return (
    <DialogFrame
      title="Danh sách gán cột"
      actions={[
        {
          text: "Lưu",
          primary: true,
          onClick: () => {
            reply({ action: "save", dateColumn, remove });
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
        <strong>Vùng nguồn:</strong> {payload.source} · {payload.dayCount} ngày
        {payload.currentDay ? ` · đang hiển thị ${payload.currentDay}` : ""}
      </Text>
      <ColumnDropdown label="Cột Ngày" options={payload.options} selected={dateColumn} onChange={setDateColumn} />
      {payload.mappings.length === 0 ? (
        <Text>Chưa gán ô nào. Chọn ô trên biểu mẫu rồi bấm &quot;Gán cột&quot; trên Ribbon.</Text>
      ) : (
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
          <thead>
            <tr>
              <th style={cellStyle}>Ô trên biểu mẫu</th>
              <th style={cellStyle}>Cột nguồn</th>
              <th style={cellStyle}>Bỏ gán</th>
            </tr>
          </thead>
          <tbody>
            {payload.mappings.map((mapping) => (
              <tr key={mapping.key}>
                <td style={cellStyle}>{mapping.target}</td>
                <td style={cellStyle}>{labelOf(mapping.column)}</td>
                <td style={cellStyle}>
                  <Checkbox
                    ariaLabel={`Bỏ gán ${mapping.target}`}
                    checked={remove.includes(mapping.key)}
                    onChange={(_event, checked) => {
                      setRemove((previous) =>
                        checked ? [...previous, mapping.key] : previous.filter((key) => key !== mapping.key),
                      );
                    }}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </DialogFrame>
  );
}
