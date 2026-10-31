import {
  DefaultButton,
  Dropdown,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
} from "@fluentui/react";
import type { JSX } from "react";
import { columnLetter } from "../../../core/address/address";
import { formatTarget } from "../../../core/mapping/types";
import { DayNavigator } from "./DayNavigator";
import { MappingList } from "./MappingList";
import { useTemplateMapping } from "./useTemplateMapping";

const MAX_ISSUES_SHOWN = 5;

/** Tab "Form mẫu" — F2: gán cột nguồn vào ô của form mẫu và điều hướng theo ngày. */
export function TemplatePanel(): JSX.Element {
  const state = useTemplateMapping();
  const { config, snapshot, headers, dayIndex, mappingErrors, currentDay, position, pending, busy, status } = state;
  const firstColumnIndex = snapshot?.firstColumnIndex ?? 0;
  const issues = dayIndex.issues;
  // Khóa chỉnh sửa nguồn/gán cột trong lúc chờ xác nhận ghi đè: xác nhận chỉ áp dụng cho bộ ô đích đã hiển thị.
  const editLocked = busy || pending !== undefined;
  const canExport =
    config !== undefined && snapshot !== undefined && dayIndex.days.length > 0 && mappingErrors.length === 0;

  return (
    <Stack tokens={{ childrenGap: 12, padding: "12px 0" }}>
      <Text variant="large">Xuất theo form mẫu</Text>

      <Stack tokens={{ childrenGap: 6 }}>
        <Text variant="mediumPlus">1. Nguồn dữ liệu (Input)</Text>
        <Text variant="small">
          Chọn vùng dữ liệu theo ngày (gồm dòng tiêu đề, ví dụ sheet NK_TheoNgay đã thêm cột phụ), rồi nhấn nút dưới.
        </Text>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Lấy vùng đang chọn"
            iconProps={{ iconName: "SelectAll" }}
            onClick={state.pickSource}
            disabled={editLocked}
          />
          {busy && <Spinner size={SpinnerSize.small} />}
        </Stack>
        {config && (
          <Text>
            <strong>Vùng nguồn:</strong> {config.source.sheetName}!{config.source.address} · {headers.length} cột ·{" "}
            {dayIndex.days.length} ngày
          </Text>
        )}
        {config && snapshot && (
          <Dropdown
            label="Cột Ngày"
            options={headers.map((header, index) => ({
              key: index,
              text: `${columnLetter(firstColumnIndex + index)} – ${header !== "" ? header : "(không có tiêu đề)"}`,
            }))}
            selectedKey={config.dateColumn}
            onChange={(_event, option) => {
              if (option) state.setDateColumn(Number(option.key));
            }}
            disabled={editLocked}
          />
        )}
        {issues.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            {issues.slice(0, MAX_ISSUES_SHOWN).map((issue) => (
              <div key={issue.rowNumber}>
                Dòng {issue.rowNumber}: {issue.reason}
              </div>
            ))}
            {issues.length > MAX_ISSUES_SHOWN && <div>… và {issues.length - MAX_ISSUES_SHOWN} dòng khác.</div>}
          </MessageBar>
        )}
      </Stack>

      {config && snapshot && (
        <Stack tokens={{ childrenGap: 6 }}>
          <Text variant="mediumPlus">2. Gán cột vào ô của form (Output)</Text>
          <Text variant="small">
            Chọn ô (hoặc ô merge) trên form mẫu, rồi nhấn &quot;Ô đang chọn&quot; ở cột tương ứng.
          </Text>
          <MappingList
            headers={headers}
            firstColumnIndex={firstColumnIndex}
            dateColumn={config.dateColumn}
            mappings={config.mappings}
            disabled={editLocked}
            onAssign={state.assignSelectedCell}
            onRemove={state.removeTarget}
          />
        </Stack>
      )}

      {config && snapshot && (
        <Stack tokens={{ childrenGap: 6 }}>
          <Text variant="mediumPlus">3. Xuất và chuyển ngày</Text>
          {mappingErrors.length > 0 && config.mappings.length > 0 && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline>
              {mappingErrors.join(" ")}
            </MessageBar>
          )}
          {pending && (
            <MessageBar
              messageBarType={MessageBarType.severeWarning}
              isMultiline
              actions={
                <div>
                  <PrimaryButton text="Ghi đè" onClick={state.confirmPending} disabled={busy} />
                  <DefaultButton text="Hủy" onClick={state.cancelPending} disabled={busy} />
                </div>
              }
            >
              <div>Các ô đích sau đang có dữ liệu và sẽ bị ghi đè:</div>
              <ul style={{ margin: "4px 0", paddingLeft: 18 }}>
                {pending.risky.map((item) => (
                  <li key={formatTarget(item.target)}>
                    {formatTarget(item.target)}
                    {item.hasFormula ? " — đang chứa CÔNG THỨC (sẽ bị thay bằng giá trị)" : ""}
                  </li>
                ))}
              </ul>
            </MessageBar>
          )}
          <PrimaryButton
            text={position === -1 ? "Xuất lên form" : "Xuất lại ngày đang chọn"}
            iconProps={{ iconName: "Export" }}
            onClick={state.exportFirst}
            disabled={busy || !canExport || pending !== undefined}
          />
          {currentDay !== undefined && position !== -1 && (
            <DayNavigator
              currentDay={currentDay}
              position={position}
              total={dayIndex.days.length}
              disabled={busy || pending !== undefined}
              onStep={state.step}
              onJump={state.jumpTo}
            />
          )}
        </Stack>
      )}

      {status && (
        <MessageBar messageBarType={status.type} isMultiline>
          {status.text}
        </MessageBar>
      )}
    </Stack>
  );
}
