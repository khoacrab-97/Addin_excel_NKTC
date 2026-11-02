import { ChoiceGroup, MessageBar, MessageBarType, Text, TextField } from "@fluentui/react";
import { type JSX, useState } from "react";
import { normalizeCellAddress } from "../../core/address/address";
import type { DialogPayloads, DialogReplies } from "../../core/dialog/protocol";
import { INPUT_TABLE_HEADERS, buildInputTableLayout, parseRowCount } from "../../core/inputTable/layout";
import { DialogFrame } from "../components/DialogFrame";

type Target = "current" | "new";

interface Props {
  payload: DialogPayloads["create-table"];
  reply: (value: DialogReplies["create-table"]) => void;
}

/** F4: chọn nơi tạo bảng nhập tiến độ, vị trí và số dòng. */
export function CreateTableView({ payload, reply }: Props): JSX.Element {
  const [target, setTarget] = useState<Target>("new");
  const [cell, setCell] = useState(payload.cell);
  const [rowsText, setRowsText] = useState(String(payload.defaultRows));

  const rows = parseRowCount(rowsText);
  const cellValid = target === "new" || normalizeCellAddress(cell) !== undefined;
  const layout =
    rows !== undefined && cellValid ? buildInputTableLayout(target === "new" ? "A1" : cell, rows) : undefined;
  const submit = (): void => {
    if (rows === undefined || layout?.ok !== true) return;
    reply({ action: "create", target, cell: target === "new" ? "A1" : (normalizeCellAddress(cell) ?? cell), rows });
  };

  return (
    <DialogFrame
      title="Tạo bảng nhập tiến độ"
      actions={[
        { text: "Tạo bảng", primary: true, disabled: layout?.ok !== true, onClick: submit },
        {
          text: "Hủy",
          onClick: () => {
            reply({ action: "cancel" });
          },
        },
      ]}
    >
      <Text>Các cột: {INPUT_TABLE_HEADERS.join(" | ")}. STT được đánh sẵn; ngày nhập dạng dd/mm/yyyy.</Text>
      <ChoiceGroup
        label="Tạo bảng trên"
        selectedKey={target}
        options={[
          { key: "new", text: "Sheet mới (TienDo), bắt đầu tại ô A1" },
          { key: "current", text: `Sheet hiện tại (${payload.sheetName})` },
        ]}
        onChange={(_event, option) => {
          if (option?.key === "new" || option?.key === "current") setTarget(option.key);
        }}
      />
      {target === "current" && (
        <TextField
          label="Ô bắt đầu (góc trên-trái, ô tiêu đề STT)"
          value={cell}
          onChange={(_event, value) => {
            setCell(value ?? "");
          }}
          errorMessage={cellValid ? "" : "Địa chỉ ô không hợp lệ (ví dụ B3)."}
        />
      )}
      <TextField
        label={`Số dòng (1–${payload.maxRows}, mặc định ${payload.defaultRows})`}
        value={rowsText}
        onChange={(_event, value) => {
          setRowsText(value ?? "");
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter") submit();
        }}
        errorMessage={rows === undefined ? `Nhập số nguyên từ 1 đến ${payload.maxRows}.` : ""}
      />
      {layout?.ok === true && (
        <Text variant="small">
          Bảng sẽ chiếm vùng {layout.layout.address}
          {target === "current" ? " — vùng này phải đang trống (add-in không ghi đè dữ liệu)." : "."}
        </Text>
      )}
      {layout?.ok === false && <MessageBar messageBarType={MessageBarType.error}>{layout.error}</MessageBar>}
    </DialogFrame>
  );
}
