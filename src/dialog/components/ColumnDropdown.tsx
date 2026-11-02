import { Dropdown } from "@fluentui/react";
import type { JSX } from "react";
import type { ColumnOption } from "../../core/dialog/protocol";

interface ColumnDropdownProps {
  label: string;
  options: readonly ColumnOption[];
  selected: number | undefined;
  onChange: (column: number) => void;
}

/** Chọn một cột của vùng dữ liệu. */
export function ColumnDropdown({ label, options, selected, onChange }: ColumnDropdownProps): JSX.Element {
  return (
    <Dropdown
      label={label}
      placeholder="Chọn cột"
      options={options.map((option) => ({ key: option.index, text: option.label }))}
      selectedKey={selected ?? null}
      onChange={(_event, option) => {
        if (option) onChange(Number(option.key));
      }}
    />
  );
}
