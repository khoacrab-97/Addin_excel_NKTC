import { MessageBarType } from "@fluentui/react";
import { useEffect, useMemo, useRef, useState } from "react";
import { sameHeaders } from "../../../core/dailyLog/prepareDailyLog";
import { formatDay, parseDateText } from "../../../core/dates/excelDate";
import { headerTexts } from "../../../core/mapping/config";
import { type DayIndex, buildDayIndex, detectDateColumn, indexOfDay, stepIndex } from "../../../core/mapping/days";
import { type MappingConfig, type TargetCell, formatTarget, targetKey } from "../../../core/mapping/types";
import { validateMappings } from "../../../core/mapping/validate";
import { buildWritePlan } from "../../../core/mapping/writePlan";
import { toUserMessage } from "../../../excel/errors";
import { type RangeSnapshot, getSelectedCell, getSelectedRangeRef, readRange } from "../../../excel/rangeReader";
import { loadMappingConfig, saveMappingConfig } from "../../../excel/settingsStore";
import { type TargetState, inspectTargets, writeCells } from "../../../excel/templateWriter";

export interface StatusMessage {
  type: MessageBarType;
  text: string;
}

/** Yêu cầu xác nhận trước khi ghi vào ô đích đang có công thức/dữ liệu (R2.4, R2.5). */
export interface PendingConfirmation {
  risky: TargetState[];
  day: number;
}

const EMPTY_INDEX: DayIndex = { days: [], issues: [] };

function mappingKey(config: MappingConfig | undefined): string {
  if (!config) return "";
  return config.mappings
    .map((mapping) => targetKey(mapping.target))
    .sort()
    .join("|");
}

/** Toàn bộ trạng thái và thao tác của tab "Form mẫu" (F2). */
export function useTemplateMapping() {
  // Cấu hình đã lưu trong workbook (R2.8), đọc một lần khi mở tab.
  const [saved] = useState(loadMappingConfig);
  const [config, setConfig] = useState<MappingConfig | undefined>(saved);
  const [snapshot, setSnapshot] = useState<RangeSnapshot>();
  const [currentDay, setCurrentDay] = useState<number | undefined>(saved?.currentDay);
  const [pending, setPending] = useState<PendingConfirmation>();
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<StatusMessage>();
  /** Bộ ô đích đã được người dùng xác nhận ghi trong phiên làm việc hiện tại. */
  const confirmedKey = useRef("");

  const dayIndex = useMemo(
    () =>
      snapshot && config ? buildDayIndex(snapshot.values, config.dateColumn, snapshot.firstRowNumber) : EMPTY_INDEX,
    [snapshot, config],
  );
  const headers = snapshot ? headerTexts(snapshot.values[0] ?? []) : (config?.headers ?? []);
  const mappingErrors = config ? validateMappings(config.mappings, headers.length) : [];

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

  const persist = async (next: MappingConfig): Promise<void> => {
    setConfig(next);
    await saveMappingConfig(next);
  };

  // Đọc lại vùng nguồn của cấu hình đã lưu; không tự ghi gì lên form.
  useEffect(() => {
    if (!saved) return;
    readRange(saved.source)
      .then((fresh) => {
        setSnapshot(fresh);
        if (!sameHeaders(headerTexts(fresh.values[0] ?? []), saved.headers)) {
          setStatus({
            type: MessageBarType.warning,
            text: "Dòng tiêu đề của vùng nguồn đã thay đổi so với cấu hình đã lưu. Vui lòng kiểm tra lại gán cột.",
          });
        }
      })
      .catch((error: unknown) => {
        setStatus({ type: MessageBarType.error, text: toUserMessage(error) });
      });
  }, [saved]);

  const pickSource = (): void => {
    void run(async () => {
      const fresh = await readRange(await getSelectedRangeRef());
      const nextHeaders = headerTexts(fresh.values[0] ?? []);
      const keepMappings = config !== undefined && sameHeaders(nextHeaders, config.headers);
      setSnapshot(fresh);
      setCurrentDay(undefined);
      await persist({
        version: 1,
        source: fresh.ref,
        headers: nextHeaders,
        dateColumn: keepMappings ? config.dateColumn : (detectDateColumn(fresh.values) ?? 0),
        mappings: keepMappings ? config.mappings : [],
      });
    });
  };

  const setDateColumn = (column: number): void => {
    if (!config) return;
    void run(() => persist({ ...config, dateColumn: column }));
  };

  const assignSelectedCell = (column: number): void => {
    if (!config) return;
    void run(async () => {
      const target = await getSelectedCell();
      const owner = config.mappings.find((mapping) => targetKey(mapping.target) === targetKey(target));
      if (owner) {
        setStatus({
          type: MessageBarType.warning,
          text: `Ô ${formatTarget(target)} đã được gán cho cột "${headers[owner.column] ?? owner.column + 1}".`,
        });
        return;
      }
      await persist({ ...config, mappings: [...config.mappings, { column, target }] });
    });
  };

  const removeTarget = (target: TargetCell): void => {
    if (!config) return;
    void run(() =>
      persist({
        ...config,
        mappings: config.mappings.filter((mapping) => targetKey(mapping.target) !== targetKey(target)),
      }),
    );
  };

  /** Ghi dữ liệu của `day` (hoặc ngày gần nhất còn tồn tại) lên form; đọc lại vùng nguồn trước (R2.7). */
  const showDay = (day: number | undefined, confirmed = false): void => {
    if (!config) return;
    void run(async () => {
      const fresh = await readRange(config.source);
      setSnapshot(fresh);
      if (!sameHeaders(headerTexts(fresh.values[0] ?? []), config.headers)) {
        setStatus({
          type: MessageBarType.warning,
          text: "Dòng tiêu đề của vùng nguồn đã thay đổi. Vui lòng kiểm tra lại gán cột rồi thực hiện lại.",
        });
        return;
      }
      const errors = validateMappings(config.mappings, config.headers.length);
      if (errors.length > 0) {
        setStatus({ type: MessageBarType.error, text: errors.join(" ") });
        return;
      }
      const { days } = buildDayIndex(fresh.values, config.dateColumn, fresh.firstRowNumber);
      if (days.length === 0) {
        setStatus({ type: MessageBarType.error, text: "Vùng nguồn không có ngày hợp lệ ở cột Ngày đã chọn." });
        return;
      }
      const found = day === undefined ? 0 : indexOfDay(days, day);
      const target = days[found === -1 ? 0 : found];
      if (!target) return;

      const key = mappingKey(config);
      if (!confirmed && confirmedKey.current !== key) {
        const states = await inspectTargets(config.mappings.map((mapping) => mapping.target));
        const risky = states.filter((state) => state.hasFormula || state.hasValue);
        if (risky.length > 0) {
          setPending({ risky, day: target.day });
          return;
        }
      }

      await writeCells(buildWritePlan(fresh.values[target.valueRowIndex] ?? [], config.mappings));
      confirmedKey.current = key;
      setPending(undefined);
      setCurrentDay(target.day);
      await persist({ ...config, currentDay: target.day });
      if (found === -1 && day !== undefined) {
        setStatus({
          type: MessageBarType.warning,
          text: `Không còn ngày ${formatDay(day)} trong vùng nguồn; đã hiển thị ngày ${formatDay(target.day)}.`,
        });
      }
    });
  };

  const position = currentDay === undefined ? -1 : indexOfDay(dayIndex.days, currentDay);

  const step = (delta: number): void => {
    const next = dayIndex.days[stepIndex(dayIndex.days.length, position, delta)];
    if (next) showDay(next.day);
  };

  const jumpTo = (text: string): void => {
    const day = parseDateText(text);
    if (day === undefined) {
      setStatus({ type: MessageBarType.error, text: "Ngày nhập không hợp lệ (định dạng dd/mm/yyyy)." });
      return;
    }
    if (indexOfDay(dayIndex.days, day) === -1) {
      setStatus({ type: MessageBarType.error, text: `Ngày ${formatDay(day)} không có trong vùng nguồn.` });
      return;
    }
    showDay(day);
  };

  return {
    config,
    snapshot,
    headers,
    dayIndex,
    mappingErrors,
    currentDay,
    position,
    pending,
    busy,
    status,
    pickSource,
    setDateColumn,
    assignSelectedCell,
    removeTarget,
    exportFirst: () => {
      showDay(currentDay);
    },
    confirmPending: () => {
      if (pending) showDay(pending.day, true);
    },
    cancelPending: () => {
      setPending(undefined);
    },
    step,
    jumpTo,
  };
}
