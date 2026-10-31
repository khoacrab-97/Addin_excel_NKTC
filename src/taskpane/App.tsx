import type { JSX } from "react";
import { Pivot, PivotItem, Stack } from "@fluentui/react";
import { DailyLogPanel } from "./components/dailyLog/DailyLogPanel";
import { TemplatePanel } from "./components/template/TemplatePanel";
import { UtilitiesPanel } from "./components/UtilitiesPanel";

const containerTokens = { padding: 12 };

export function App(): JSX.Element {
  return (
    <Stack tokens={containerTokens}>
      <Pivot aria-label="Chức năng nhật ký thi công">
        <PivotItem headerText="Theo ngày" itemKey="daily">
          <DailyLogPanel />
        </PivotItem>
        <PivotItem headerText="Form mẫu" itemKey="template">
          <TemplatePanel />
        </PivotItem>
        <PivotItem headerText="Tiện ích" itemKey="utilities">
          <UtilitiesPanel />
        </PivotItem>
      </Pivot>
    </Stack>
  );
}
