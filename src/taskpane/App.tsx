import { DefaultButton, MessageBar, MessageBarType, Pivot, PivotItem, Stack } from "@fluentui/react";
import { type JSX, useState } from "react";
import { UtilitiesPanel } from "./components/UtilitiesPanel";
import { DailyLogPanel } from "./components/dailyLog/DailyLogPanel";
import { TemplatePanel } from "./components/template/TemplatePanel";
import { useUpdateCheck } from "./components/utilities/useUpdateCheck";

const containerTokens = { padding: 12 };
type TabKey = "daily" | "template" | "utilities";

export function App(): JSX.Element {
  const [tab, setTab] = useState<TabKey>("daily");
  const update = useUpdateCheck();

  return (
    <Stack tokens={containerTokens}>
      {update.status?.kind === "available" && tab !== "utilities" && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline={false}
          actions={
            <DefaultButton
              text="Xem hướng dẫn cập nhật"
              onClick={() => {
                setTab("utilities");
              }}
            />
          }
        >
          Đã có phiên bản mới {update.status.latest}.
        </MessageBar>
      )}
      <Pivot
        aria-label="Chức năng nhật ký thi công"
        selectedKey={tab}
        onLinkClick={(item) => {
          const key = item?.props.itemKey;
          if (key === "daily" || key === "template" || key === "utilities") setTab(key);
        }}
      >
        <PivotItem headerText="Theo ngày" itemKey="daily">
          <DailyLogPanel />
        </PivotItem>
        <PivotItem headerText="Form mẫu" itemKey="template">
          <TemplatePanel />
        </PivotItem>
        <PivotItem headerText="Tiện ích" itemKey="utilities">
          <UtilitiesPanel updateStatus={update.status} checking={update.checking} onCheckUpdate={update.recheck} />
        </PivotItem>
      </Pivot>
    </Stack>
  );
}
