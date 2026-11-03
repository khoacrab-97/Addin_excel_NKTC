import { stripSheetName } from "../core/address/address";
import type { RangeRef } from "./rangeReader";

/**
 * Viền nháy quanh vùng/ô đang chọn trong lúc mở thanh chọn (giống viền "kiến bò" khi Copy của Excel).
 * Office.js không gọi được hiệu ứng Copy, nên add-in vẽ tạm một hình chữ nhật không nền, viền nét đứt
 * (Shape, ExcelApi 1.9) đè lên vùng, đổi màu viền liên tục, rồi xóa khi đóng thanh chọn.
 * Không đụng tới giá trị, công thức, định dạng ô. Mọi lỗi (sheet bị khóa, Excel cũ…) đều bỏ qua: không có
 * viền nháy thì việc chọn vẫn chạy bình thường.
 */

/** Tên hình chữ nhật tạm — dùng để dọn hình còn sót (ví dụ Excel bị đóng giữa chừng). */
export const FLASH_SHAPE_NAME = "NKTC_VienChon";
const COLORS = ["#107C41", "#FFFFFF"] as const;
const BLINK_MS = 450;
const LINE_WEIGHT = 2.25;
/** Lùi viền ra ngoài vùng một chút để không che nội dung ô. */
const OUTSET = 1.5;

export function canFlashSelection(): boolean {
  return Office.context.requirements.isSetSupported("ExcelApi", "1.9");
}

/** Xóa mọi hình viền nháy còn sót trên các sheet. */
export async function clearFlashLeftovers(): Promise<void> {
  if (!canFlashSelection()) return;
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const collections = sheets.items.map((sheet) => {
        const shapes = sheet.shapes;
        shapes.load("items/name");
        return shapes;
      });
      await context.sync();
      for (const shapes of collections) {
        for (const shape of shapes.items) if (shape.name === FLASH_SHAPE_NAME) shape.delete();
      }
      await context.sync();
    });
  } catch {
    // Không dọn được (sheet bị khóa…) → bỏ qua.
  }
}

interface DrawnShape {
  sheetName: string;
  id: string;
}

/** Một phiên viền nháy (một lần mở thanh chọn). Các thao tác Excel được xếp hàng, không chạy chồng nhau. */
export class SelectionFlash {
  private queue: Promise<void> = Promise.resolve();
  private shape: DrawnShape | undefined;
  private timer: number | undefined;
  private tick = 0;
  private busy = false;
  private stopped = false;

  constructor() {
    // Dọn viền còn sót từ lần trước trước khi vẽ viền mới (cùng hàng đợi → không xóa nhầm viền mới).
    this.enqueue(clearFlashLeftovers);
  }

  /** Vẽ viền quanh `ref` (thay viền cũ). `expandMerge`: ô nằm trong vùng gộp → viền quanh cả vùng gộp. */
  show(ref: RangeRef, expandMerge = false): void {
    if (this.stopped) return;
    this.enqueue(async () => {
      await this.remove();
      if (this.stopped) return;
      this.shape = await drawFrame(ref, expandMerge);
    });
    this.timer ??= window.setInterval(() => {
      this.blink();
    }, BLINK_MS);
  }

  /** Dừng nháy và xóa viền. */
  async stop(): Promise<void> {
    this.stopped = true;
    if (this.timer !== undefined) window.clearInterval(this.timer);
    this.timer = undefined;
    this.enqueue(() => this.remove());
    await this.queue;
  }

  private blink(): void {
    // Lần đổi màu trước chưa xong → bỏ lượt này (không dồn hàng đợi).
    if (this.busy || !this.shape || this.stopped) return;
    const shape = this.shape;
    this.tick += 1;
    const color = COLORS[this.tick % COLORS.length] ?? COLORS[0];
    this.enqueue(() => setLineColor(shape, color));
  }

  private enqueue(task: () => Promise<void>): void {
    this.queue = this.queue.then(async () => {
      this.busy = true;
      try {
        await task();
      } catch {
        // Lỗi vẽ/xóa viền không ảnh hưởng việc chọn.
      } finally {
        this.busy = false;
      }
    });
  }

  private async remove(): Promise<void> {
    const shape = this.shape;
    this.shape = undefined;
    if (!shape) return;
    await Excel.run(async (context) => {
      const item = context.workbook.worksheets.getItem(shape.sheetName).shapes.getItemOrNullObject(shape.id);
      await context.sync();
      if (!item.isNullObject) item.delete();
      await context.sync();
    });
  }
}

async function drawFrame(ref: RangeRef, expandMerge: boolean): Promise<DrawnShape> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(ref.sheetName);
    let range = sheet.getRange(ref.address);
    if (expandMerge && Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
      const merged = range.getMergedAreasOrNullObject();
      merged.load("address");
      await context.sync();
      if (!merged.isNullObject && !merged.address.includes(",")) range = sheet.getRange(stripSheetName(merged.address));
    }
    range.load(["left", "top", "width", "height"]);
    await context.sync();
    const shape = sheet.shapes.addGeometricShape("Rectangle");
    shape.name = FLASH_SHAPE_NAME;
    // Sát cột A / dòng 1 thì không lùi ra ngoài được (tọa độ âm) → giữ mép phải/dưới đúng chỗ.
    const left = Math.max(0, range.left - OUTSET);
    const top = Math.max(0, range.top - OUTSET);
    shape.left = left;
    shape.top = top;
    shape.width = range.left + range.width + OUTSET - left;
    shape.height = range.top + range.height + OUTSET - top;
    shape.fill.clear();
    shape.lineFormat.color = COLORS[0];
    shape.lineFormat.weight = LINE_WEIGHT;
    shape.lineFormat.dashStyle = "Dash";
    shape.load("id");
    await context.sync();
    return { sheetName: ref.sheetName, id: shape.id };
  });
}

async function setLineColor(shape: DrawnShape, color: string): Promise<void> {
  await Excel.run(async (context) => {
    const item = context.workbook.worksheets.getItem(shape.sheetName).shapes.getItemOrNullObject(shape.id);
    await context.sync();
    if (item.isNullObject) return;
    item.lineFormat.color = color;
    await context.sync();
  });
}
