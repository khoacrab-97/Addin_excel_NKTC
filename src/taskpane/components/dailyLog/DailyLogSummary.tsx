import { MessageBar, MessageBarType, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import type { PrepareDailyLogResult } from "../../../core/dailyLog/prepareDailyLog";
import { formatDay } from "../../../core/dates/excelDate";

const MAX_ISSUES_SHOWN = 10;
const PREVIEW_DAYS = 5;

const cellStyle = { border: "1px solid #c8c8c8", padding: "4px 6px", verticalAlign: "top" } as const;

interface DailyLogSummaryProps {
  preparation: PrepareDailyLogResult;
}

/** Tóm tắt kiểm tra dữ liệu và xem trước kết quả F1. */
export function DailyLogSummary({ preparation }: DailyLogSummaryProps): JSX.Element {
  const issues = preparation.parsed?.issues ?? [];
  const itemCount = preparation.parsed?.items.length ?? 0;

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {preparation.ok ? (
        <MessageBar messageBarType={MessageBarType.info} isMultiline>
          {itemCount} công việc hợp lệ · từ {formatDay(preparation.built.firstDay)} đến{" "}
          {formatDay(preparation.built.lastDay)} · {preparation.built.entries.length} ngày
        </MessageBar>
      ) : (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {preparation.error}
        </MessageBar>
      )}

      {issues.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <div>{issues.length} dòng lỗi sẽ bị bỏ qua khi xuất:</div>
          <ul style={{ margin: "4px 0", paddingLeft: 18 }}>
            {issues.slice(0, MAX_ISSUES_SHOWN).map((issue) => (
              <li key={issue.rowNumber}>
                Dòng {issue.rowNumber}: {issue.reason}
              </li>
            ))}
          </ul>
          {issues.length > MAX_ISSUES_SHOWN && <div>… và {issues.length - MAX_ISSUES_SHOWN} dòng khác.</div>}
        </MessageBar>
      )}

      {preparation.ok && (
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="mediumPlus">
            Xem trước ({Math.min(PREVIEW_DAYS, preparation.built.entries.length)} ngày đầu)
          </Text>
          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ ...cellStyle, width: 80 }}>Ngày</th>
                <th style={cellStyle}>Nội dung</th>
              </tr>
            </thead>
            <tbody>
              {preparation.built.entries.slice(0, PREVIEW_DAYS).map((entry) => (
                <tr key={entry.day}>
                  <td style={cellStyle}>{formatDay(entry.day)}</td>
                  <td style={{ ...cellStyle, whiteSpace: "pre-wrap" }}>{entry.content}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Stack>
      )}
    </Stack>
  );
}
