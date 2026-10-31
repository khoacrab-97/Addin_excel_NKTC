import { ActionButton, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import { columnLetter } from "../../../core/address/address";
import { type ColumnMapping, type TargetCell, formatTarget } from "../../../core/mapping/types";

interface MappingListProps {
  headers: readonly string[];
  firstColumnIndex: number;
  dateColumn: number;
  mappings: readonly ColumnMapping[];
  disabled: boolean;
  onAssign: (column: number) => void;
  onRemove: (target: TargetCell) => void;
}

const rowStyle = { borderBottom: "1px solid #edebe9", padding: "4px 0" } as const;

/** Danh sách cột nguồn và các ô đích được gán (bên Output). */
export function MappingList(props: MappingListProps): JSX.Element {
  const { headers, firstColumnIndex, dateColumn, mappings, disabled, onAssign, onRemove } = props;
  return (
    <Stack>
      {headers.map((header, column) => {
        const letter = columnLetter(firstColumnIndex + column);
        const targets = mappings.filter((mapping) => mapping.column === column).map((mapping) => mapping.target);
        return (
          <Stack key={column} style={rowStyle} tokens={{ childrenGap: 2 }}>
            <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
              <Text styles={{ root: { fontWeight: 600 } }}>
                {letter} – {header !== "" ? header : "(không có tiêu đề)"}
                {column === dateColumn ? " (cột Ngày)" : ""}
              </Text>
              <ActionButton
                iconProps={{ iconName: "Add" }}
                text="Ô đang chọn"
                ariaLabel={`Gán ô đang chọn cho cột ${letter}`}
                onClick={() => {
                  onAssign(column);
                }}
                disabled={disabled}
              />
            </Stack>
            {targets.map((target) => (
              <Stack key={formatTarget(target)} horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                <Text variant="small">→ {formatTarget(target)}</Text>
                <ActionButton
                  text="× Bỏ gán"
                  ariaLabel={`Bỏ gán ô ${formatTarget(target)}`}
                  title="Bỏ gán"
                  onClick={() => {
                    onRemove(target);
                  }}
                  disabled={disabled}
                />
              </Stack>
            ))}
          </Stack>
        );
      })}
    </Stack>
  );
}
