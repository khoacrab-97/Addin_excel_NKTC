import { DefaultButton, Dropdown, MessageBar, MessageBarType, Stack, Text } from "@fluentui/react";
import { type JSX, useState } from "react";
import type { DialogPayloads, DialogReplies, ExportFormState, FormRowState } from "../../core/dialog/protocol";
import { ColumnDropdown } from "../components/ColumnDropdown";
import { DialogFrame } from "../components/DialogFrame";
import { RangeField } from "../components/RangeField";

const MAX_ROWS = 500;

interface Props {
  payload: DialogPayloads["export-form"];
  reply: (value: DialogReplies["export-form"]) => void;
}

const cellStyle = { padding: "4px 6px", verticalAlign: "top" } as const;

/**
 * F2 "Xuất vào form": vùng nguồn (dòng đầu là tiêu đề) + các dòng gán "cột nguồn (trái) → ô trên form (phải)".
 * + Thêm để thêm dòng; chọn dòng rồi Xóa dòng (hoặc phím Delete); Gán để ghi ngày đầu tiên lên form.
 */
export function ExportFormView({ payload, reply }: Props): JSX.Element {
  const [rows, setRows] = useState<FormRowState[]>(
    payload.state.rows.length > 0 ? payload.state.rows : [{ target: "" }],
  );
  const [selected, setSelected] = useState<number | undefined>(payload.state.selected);
  const [dateColumn, setDateColumn] = useState<number | undefined>(payload.state.dateColumn);
  /** Tăng khi xóa dòng → các ô nhập được dựng lại đúng theo dòng mới. */
  const [version, setVersion] = useState(0);
  const sourceOk = payload.source.error === undefined && payload.options.length > 0;

  const state = (): ExportFormState => {
    const next: ExportFormState = { rows };
    if (dateColumn !== undefined) next.dateColumn = dateColumn;
    if (selected !== undefined) next.selected = selected;
    return next;
  };
  const updateRow = (index: number, change: Partial<FormRowState>): void => {
    setRows((previous) => previous.map((row, i) => (i === index ? { ...row, ...change } : row)));
  };
  const removeSelected = (): void => {
    if (selected === undefined) return;
    setRows((previous) => {
      const next = previous.filter((_row, i) => i !== selected);
      return next.length > 0 ? next : [{ target: "" }];
    });
    setSelected(undefined);
    setVersion((value) => value + 1);
  };

  return (
    <DialogFrame
      title="Xuất vào form"
      actions={[
        {
          text: "Gán",
          primary: true,
          disabled: !sourceOk,
          onClick: () => {
            reply({ action: "assign", state: state() });
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
        label="Vùng nguồn (dòng đầu là tiêu đề, gồm cả các cột đã thêm)"
        ariaLabel="Vùng nguồn"
        value={payload.source.value}
        error={payload.source.error}
        onPick={(address) => {
          reply(
            address === undefined
              ? { action: "pick-source", state: state() }
              : { action: "pick-source", address, state: state() },
          );
        }}
      />
      {sourceOk && (
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 12 }}>
          <Stack.Item grow>
            <ColumnDropdown
              label="Cột ngày (dùng để chuyển ngày)"
              options={payload.options}
              selected={dateColumn}
              onChange={setDateColumn}
            />
          </Stack.Item>
          <Text styles={{ root: { paddingBottom: 6 } }}>{payload.dayCount} ngày</Text>
        </Stack>
      )}

      {sourceOk && (
        <div
          role="grid"
          aria-label="Danh sách gán"
          tabIndex={-1}
          onKeyDown={(event) => {
            const target = event.target as HTMLElement;
            if (event.key === "Delete" && target.tagName !== "INPUT") removeSelected();
          }}
        >
          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
            <thead>
              <tr>
                <th style={{ ...cellStyle, width: 28 }}>#</th>
                <th style={{ ...cellStyle, textAlign: "left" }}>Cột nguồn</th>
                <th style={{ ...cellStyle, textAlign: "left" }}>Ô trên form</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, index) => (
                <tr
                  key={index}
                  aria-selected={selected === index}
                  onClick={() => {
                    setSelected(index);
                  }}
                  style={{ background: selected === index ? "#deecf9" : undefined, cursor: "pointer" }}
                >
                  <td style={cellStyle}>{index + 1}</td>
                  <td style={cellStyle}>
                    <Dropdown
                      ariaLabel={`Cột nguồn dòng ${index + 1}`}
                      placeholder="Chọn cột"
                      options={payload.options.map((option) => ({ key: option.index, text: option.label }))}
                      selectedKey={row.column ?? null}
                      onChange={(_event, option) => {
                        if (option) updateRow(index, { column: Number(option.key) });
                      }}
                    />
                  </td>
                  <td style={cellStyle}>
                    <RangeField
                      key={`${version}-${index}`}
                      label=""
                      ariaLabel={`Ô trên form dòng ${index + 1}`}
                      placeholder="Mẫu NKTC!E2"
                      value={row.target}
                      error={payload.rowError?.row === index ? payload.rowError.message : undefined}
                      onTextChange={(text) => {
                        updateRow(index, { target: text });
                      }}
                      onPick={(address) => {
                        const current = state();
                        current.selected = index;
                        reply(
                          address === undefined
                            ? { action: "pick-target", row: index, state: current }
                            : { action: "pick-target", row: index, address, state: current },
                        );
                      }}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {sourceOk && (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="+ Thêm"
            disabled={rows.length >= MAX_ROWS}
            onClick={() => {
              setRows((previous) => [...previous, { target: "" }]);
              setSelected(rows.length);
            }}
          />
          <DefaultButton text="Xóa dòng" disabled={selected === undefined} onClick={removeSelected} />
        </Stack>
      )}
      {sourceOk && (
        <Text variant="small">
          Mỗi dòng: chọn cột nguồn (trái); bên phải chọn ô (hoặc ô gộp) trên biểu mẫu rồi bấm ⬆, hoặc gõ ô (ví dụ Mẫu
          NKTC!E2) rồi Enter. Bấm Gán: biểu mẫu nhận dữ liệu ngày đầu tiên.
        </Text>
      )}
      {payload.errors && payload.errors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {payload.errors.map((line) => (
            <div key={line}>{line}</div>
          ))}
        </MessageBar>
      )}
    </DialogFrame>
  );
}
