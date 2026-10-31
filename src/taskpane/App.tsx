import { Pivot, PivotItem, Stack } from "@fluentui/react";
import { PlaceholderPanel } from "./components/PlaceholderPanel";
import { UtilitiesPanel } from "./components/UtilitiesPanel";

const containerTokens = { padding: 12 };

export function App(): JSX.Element {
  return (
    <Stack tokens={containerTokens}>
      <Pivot aria-label="Chức năng nhật ký thi công">
        <PivotItem headerText="Theo ngày" itemKey="daily">
          <PlaceholderPanel title="Xuất nhật ký theo ngày" phase="Phase 1" />
        </PivotItem>
        <PivotItem headerText="Form mẫu" itemKey="template">
          <PlaceholderPanel title="Xuất theo form mẫu" phase="Phase 2" />
        </PivotItem>
        <PivotItem headerText="Tiện ích" itemKey="utilities">
          <UtilitiesPanel />
        </PivotItem>
      </Pivot>
    </Stack>
  );
}
