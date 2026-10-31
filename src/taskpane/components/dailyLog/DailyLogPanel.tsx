import {
  DefaultButton,
  Dropdown,
  type IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
} from "@fluentui/react";
import { type JSX, useMemo, useState } from "react";
import { type ColumnKey, type DailyLogColumns, detectColumns } from "../../../core/dailyLog/columns";
import { prepareDailyLog, sameHeaders } from "../../../core/dailyLog/prepareDailyLog";
import { columnLetter } from "../../../core/address/address";
import { writeDailyLogSheet } from "../../../excel/dailyLogWriter";
import { toUserMessage } from "../../../excel/errors";
import { type RangeSnapshot, getSelectedRangeRef, readRange } from "../../../excel/rangeReader";
import { DailyLogSummary } from "./DailyLogSummary";

const COLUMN_FIELDS: ReadonlyArray<{ key: ColumnKey; label: string }> = [
  { key: "content", label: "Nội dung công việc" },
  { key: "start", label: "Ngày bắt đầu" },
  { key: "end", label: "Ngày kết thúc" },
];

interface StatusMessage {
  type: MessageBarType;
  text: string;
}

/** Tab "Theo ngày" — F1: xuất nhật ký theo ngày ra sheet mới. */
export function DailyLogPanel(): JSX.Element {
  const [snapshot, setSnapshot] = useState<RangeSnapshot>();
  const [columns, setColumns] = useState<Partial<DailyLogColumns>>({});
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<StatusMessage>();

  const headers = snapshot?.values[0] ?? [];
  const preparation = useMemo(
    () => (snapshot ? prepareDailyLog(snapshot.values, columns, snapshot.firstRowNumber) : undefined),
    [snapshot, columns],
  );

  const columnOptions: IDropdownOption[] = headers.map((header, index) => ({
    key: index,
    text: `${columnLetter((snapshot?.firstColumnIndex ?? 0) + index)} – ${typeof header === "string" && header.trim() !== "" ? header : "(không có tiêu đề)"}`,
  }));

  const run = async (action: () => Promise<void>): Promise<void> => {
    setBusy(true);
    setStatus(undefined);
    try {
      await action();
    } catch (error) {
      setStatus({ type: MessageBarType.error, text: toUserMessage(error) });
    } finally {
      setBusy(false);
    }
  };

  const handlePickRange = (): void => {
    void run(async () => {
      const ref = await getSelectedRangeRef();
      const next = await readRange(ref);
      setSnapshot(next);
      setColumns(detectColumns(next.values[0] ?? []));
    });
  };

  const handleExport = (): void => {
    if (!snapshot) return;
    void run(async () => {
      // Đọc lại dữ liệu mới nhất trước khi xuất (CLAUDE.md §4).
      const fresh = await readRange(snapshot.ref);
      if (!sameHeaders(fresh.values[0] ?? [], headers) || fresh.firstColumnIndex !== snapshot.firstColumnIndex) {
        setSnapshot(fresh);
        setColumns(detectColumns(fresh.values[0] ?? []));
        setStatus({
          type: MessageBarType.warning,
          text: "Dòng tiêu đề của vùng input đã thay đổi. Vui lòng kiểm tra lại cột rồi nhấn Xuất lần nữa.",
        });
        return;
      }
      setSnapshot(fresh);
      const result = prepareDailyLog(fresh.values, columns, fresh.firstRowNumber);
      if (!result.ok) {
        setStatus({ type: MessageBarType.error, text: result.error });
        return;
      }
      const sheetName = await writeDailyLogSheet(result.built.entries);
      setStatus({
        type: MessageBarType.success,
        text: `Đã tạo sheet "${sheetName}" gồm ${result.built.entries.length} ngày.`,
      });
    });
  };

  return (
    <Stack tokens={{ childrenGap: 12, padding: "12px 0" }}>
      <Text variant="large">Xuất nhật ký theo ngày</Text>
      <Text variant="small">
        Chọn vùng danh sách công việc trên sheet (gồm dòng tiêu đề), sau đó nhấn &quot;Lấy vùng đang chọn&quot;.
      </Text>

      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <DefaultButton
          text="Lấy vùng đang chọn"
          iconProps={{ iconName: "SelectAll" }}
          onClick={handlePickRange}
          disabled={busy}
        />
        {busy && <Spinner size={SpinnerSize.small} />}
      </Stack>

      {snapshot && (
        <Text>
          <strong>Vùng input:</strong> {snapshot.ref.sheetName}!{snapshot.ref.address}
        </Text>
      )}

      {snapshot &&
        COLUMN_FIELDS.map((field) => (
          <Dropdown
            key={field.key}
            label={field.label}
            placeholder="Chọn cột"
            options={columnOptions}
            selectedKey={columns[field.key] ?? null}
            onChange={(_event, option) => {
              if (option) setColumns((previous) => ({ ...previous, [field.key]: Number(option.key) }));
            }}
            disabled={busy}
          />
        ))}

      {preparation && <DailyLogSummary preparation={preparation} />}

      {snapshot && (
        <PrimaryButton
          text="Xuất nhật ký theo ngày"
          iconProps={{ iconName: "Export" }}
          onClick={handleExport}
          disabled={busy || !preparation?.ok}
        />
      )}

      {status && (
        <MessageBar messageBarType={status.type} isMultiline>
          {status.text}
        </MessageBar>
      )}
    </Stack>
  );
}
