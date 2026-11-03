/*
 * Smoke test F2 (nhóm Form mẫu + Chuyển ngày trên Ribbon) trên bundle production với Excel giả lập:
 * Chèn form mẫu, hộp thoại chung "Xuất vào form" (vùng nguồn + dòng gán + Gán), chuyển ngày.
 * Dữ liệu: test/fixtures/demo_f2.json + demo_f2.targets.json (ô đích của từng cột).
 * Log chỉ in tên bước và số liệu, không in nội dung dữ liệu test.
 *
 *   node scripts/smoke/run-f2.js [tên file trong test/fixtures]
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { chromium } = require("playwright");
const { fixturePath: resolveFixture } = require("./fixtures");
const { openWorkbook } = require("./harness");

const fixturePath = resolveFixture(
  process.argv[2] && !process.argv[2].startsWith("--") ? process.argv[2] : "demo_f2.json",
);
const fixture = JSON.parse(fs.readFileSync(fixturePath, "utf8"));
/** Ô đích của từng cột nguồn: { column, header, cell, select } — `select` là vùng merge Excel trả về khi chọn ô. */
const meta = JSON.parse(fs.readFileSync(fixturePath.replace(/\.json$/, ".targets.json"), "utf8"));
const SHEET = meta.sheet;
const MAPPINGS = meta.mappings;
const COLUMN_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const daily = fixture.sheets.find((sheet) => sheet.name === "NK_TheoNgay");
const formSheet = fixture.sheets.find((sheet) => sheet.name === SHEET);
const lastColumn = COLUMN_LETTERS[daily.values[0].length - 1];
const sourceAddress = `A1:${lastColumn}${daily.values.length}`;
const TARGET_CELLS = new Set(MAPPINGS.map((m) => `${SHEET}!${m.cell}`));
const SETTINGS_KEY = "nktc.templateMapping.v1";

function sourceRowForDate(label) {
  const [d, m, y] = label.split("/").map(Number);
  const serial = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const row = daily.values.find((r) => Math.floor(r[0]) === serial);
  assert.ok(row, `Không có ngày ${label} trong NK_TheoNgay`);
  return row;
}

const steps = [];
function pass(name, detail = "") {
  steps.push(`  ✔ ${name}${detail ? ` — ${detail}` : ""}`);
}

async function main() {
  const browser = await chromium.launch();
  const errors = [];
  // Workbook chưa có biểu mẫu; Excel hỗ trợ ExcelApi 1.13 (chèn sheet, nhận biết ô gộp).
  const startSheets = fixture.sheets.filter((sheet) => sheet.name !== SHEET);
  let wb = await openWorkbook(
    {
      sheets: startSheets,
      templateSheets: [formSheet],
      selection: { sheet: "NK_TheoNgay", address: "B7" },
      apiMinor: 13,
    },
    { browser },
  );
  const cell = (address) => wb.smoke(([s, a]) => globalThis.__smoke.cell(s, a), [SHEET, address]);
  const select = (sheet, address) => wb.smoke(([s, a]) => globalThis.__smoke.select(s, a), [sheet, address]);
  const logLength = () => wb.smoke(() => globalThis.__smoke.log.length);
  const writesSince = (start) =>
    wb.smoke((s) => globalThis.__smoke.log.slice(s).filter((e) => e.op === "values"), start);
  const config = async () => (await wb.smoke(() => globalThis.__smoke.settings()))[SETTINGS_KEY];

  async function pressExpectMessage(name) {
    const done = await wb.press(name);
    const d = await wb.dialog();
    const text = await d.text();
    await d.click("Đóng");
    await done();
    return text;
  }
  async function pressSilently(name) {
    const count = await wb.smoke(() => globalThis.__smoke.dialogCount());
    const done = await wb.press(name);
    await done();
    assert.equal(await wb.smoke(() => globalThis.__smoke.dialogCount()), count, `${name} mở hộp thoại ngoài dự kiến`);
  }
  /** Giá trị mong đợi trên form: cột Ngày ở ô mẫu "Ngày … tháng … năm …" → chữ đã rải ngày (R2.15). */
  function expectedOnForm(m, label) {
    const value = sourceRowForDate(label)[COLUMN_LETTERS.indexOf(m.column)];
    if (m.column !== "A") return value;
    const [dd, mm, yyyy] = label.split("/");
    return `Ngày ${dd} tháng ${mm} năm ${yyyy}`;
  }
  async function assertDayOnForm(label, mappings = MAPPINGS) {
    for (const m of mappings) {
      const expected = expectedOnForm(m, label);
      const actual = (await cell(m.cell))?.value ?? "";
      assert.deepEqual(actual, expected, `${SHEET}!${m.cell} (cột ${m.column} – ${m.header}) ngày ${label}`);
    }
  }
  const shapes = () => wb.smoke((name) => globalThis.__smoke.shapes(name), SHEET);
  /** Chờ đúng 1 viền nháy (khung nét đứt, không nền) bao quanh vùng `address` của sheet biểu mẫu. */
  async function waitFrame(address) {
    const [a, b = a] = address.split(":");
    const col = (ref) => COLUMN_LETTERS.indexOf(ref.replace(/\d+/g, ""));
    const row = (ref) => Number(ref.replace(/[A-Z]+/g, "")) - 1;
    const { col: cw, row: rh } = await wb.smoke(() => globalThis.__smoke.geometry);
    const left = Math.max(0, col(a) * cw - 1.5);
    const want = { left, top: Math.max(0, row(a) * rh - 1.5), width: (col(b) + 1) * cw + 1.5 - left };
    await wb.commands.waitForFunction(
      ([sheet, w]) => {
        const list = globalThis.__smoke.shapes(sheet);
        return list.length === 1 && list[0].left === w.left && list[0].top === w.top && list[0].width === w.width;
      },
      [SHEET, want],
    );
    const [frame] = await shapes();
    assert.equal(frame.name, "NKTC_VienChon");
    assert.equal(frame.fill, "none");
    assert.equal(frame.dash, "Dash");
  }
  const sourceField = (d) => d.page.getByRole("textbox", { name: "Vùng nguồn" });
  const targetField = (d, row) => d.page.getByRole("textbox", { name: `Ô trên form dòng ${row}` });
  async function chooseColumn(d, row, letter) {
    await d.page.getByRole("combobox", { name: `Cột nguồn dòng ${row}` }).click();
    await d.page
      .getByRole("option")
      .filter({ hasText: `Cột ${letter} – ` })
      .first()
      .click();
  }

  // 0. Chưa gán → các nút Chuyển ngày chỉ hướng dẫn, không ghi gì.
  let text = await pressExpectMessage("nktcDayNext");
  assert.match(text, /Chưa gán cột vào form/);
  pass("Chưa gán", "nút Chuyển ngày chỉ hướng dẫn bấm Xuất vào form");

  // 1. Chèn form mẫu (ExcelApi 1.13): thêm sheet "Mẫu NKTC" ở cuối; bấm lần 2 → không ghi đè.
  text = await pressExpectMessage("nktcInsertTemplate");
  assert.match(text, new RegExp(`Đã chèn sheet "${SHEET}"`));
  assert.equal((await wb.smoke(() => globalThis.__smoke.sheetNames())).at(-1), SHEET);
  assert.equal(await wb.smoke(() => globalThis.__smoke.activeSheet()), SHEET);
  const inserted = (await wb.smoke(() => globalThis.__smoke.log)).find((e) => e.op === "insertWorksheets");
  assert.ok(inserted.base64Length > 1000, "base64 của Mau_NKTC.xlsx");
  text = await pressExpectMessage("nktcInsertTemplate");
  assert.match(text, /đã có sheet "Mẫu NKTC"/);
  pass("Chèn form mẫu", "thêm sheet biểu mẫu từ templates/Mau_NKTC.xlsx; lần 2 báo đã có, không ghi đè");

  // 2. Xuất vào form: vùng nguồn tự nhận diện = sheet NK_TheoNgay; cột Ngày nhận diện; 1 dòng gán trống.
  await select(SHEET, "A1");
  let done = await wb.press("nktcExportForm");
  let d = await wb.dialog();
  assert.equal(await sourceField(d).inputValue(), `NK_TheoNgay!${sourceAddress}`);
  text = await d.text();
  assert.ok(text.includes("200 ngày"), text);
  assert.ok(
    (await d.page.getByRole("combobox", { name: "Cột ngày (dùng để chuyển ngày)" }).innerText()).startsWith(
      "Cột A – Ngày",
    ),
  );
  assert.equal(await d.page.getByRole("row").count(), 2); // tiêu đề + 1 dòng gán
  pass("Mở Xuất vào form", "vùng nguồn = cả sheet NK_TheoNgay (gồm cột đã thêm), cột Ngày nhận diện, số ngày đúng");

  // 3. Gán từng dòng: chọn cột (trái), chọn ô trên form rồi bấm ⬆ (phải; ô gộp → ô trên-trái); + Thêm.
  for (let i = 0; i < MAPPINGS.length; i++) {
    const m = MAPPINGS[i];
    const row = i + 1;
    if (i > 0) await d.page.getByRole("button", { name: "+ Thêm" }).click();
    assert.equal(await targetField(d, row).inputValue(), "", "dòng mới: ô trên form để trống");
    await chooseColumn(d, row, m.column);
    // ⬆ → hộp thoại thu gọn thành thanh chọn → chọn ô trên biểu mẫu (địa chỉ hiện ngay) → ⬇ / Enter → mở lại.
    d = await wb.pickVia(d, `Chọn Ô trên form dòng ${row} trên sheet`, [SHEET, m.select], {
      initial: "",
      shown: `${SHEET}!${m.cell}`,
      confirm: i % 2 === 0 ? "arrow" : "enter",
      inspect:
        i === 0
          ? async (bar) => {
              assert.ok((await bar.text()).includes("Xuất vào form – Ô trên form (dòng 1)"), "tiêu đề thanh chọn");
              assert.ok(!(await bar.text()).includes("Enter"), "không hướng dẫn bấm Enter (Enter thuộc về Excel)");
              await bar.page.screenshot({ path: path.join("smoke-output", "f2-pick-bar.png") });
              // Viền nháy quanh cả vùng gộp vừa chọn (E2:I2), đổi màu liên tục.
              await waitFrame(m.select);
              await wb.commands.waitForFunction(() =>
                globalThis.__smoke.log.some((e) => e.op === "shape.color" && e.value === "#FFFFFF"),
              );
            }
          : undefined,
    });
    assert.equal((await shapes()).length, 0, "đóng thanh chọn → xóa viền nháy");
    assert.equal(await targetField(d, row).inputValue(), `${SHEET}!${m.cell}`);
  }
  await d.page.screenshot({ path: path.join("smoke-output", "f2-export-form.png"), fullPage: true });
  // Thanh chọn: gõ địa chỉ (không ghi sheet) → sheet biểu mẫu; đóng thanh (X) → giữ ô cũ.
  d = await wb.pickVia(d, "Chọn Ô trên form dòng 2 trên sheet", undefined, { type: "B3", confirm: "enter" });
  assert.equal(await targetField(d, 2).inputValue(), `${SHEET}!B3`);
  // Viền còn sót từ lần trước (Excel bị đóng giữa chừng) → dọn trước khi vẽ viền mới; shape khác giữ nguyên.
  await wb.smoke(([s]) => globalThis.__smoke.addShape(s, "NKTC_VienChon"), [SHEET]);
  await wb.smoke(() => globalThis.__smoke.addShape("NK_TheoNgay", "NKTC_VienChon"));
  await wb.smoke(() => globalThis.__smoke.addShape("NK_TheoNgay", "Logo"));
  let start = await logLength();
  d = await wb.pickVia(d, "Chọn Ô trên form dòng 2 trên sheet", [SHEET, MAPPINGS[1].select], {
    initial: `${SHEET}!B3`,
    shown: `${SHEET}!${MAPPINGS[1].cell}`,
    // Mở thanh chọn → nháy viền ô đang gán (B3); chọn ô khác → viền chuyển sang ô mới (chỉ 1 viền).
    beforeSelect: () => waitFrame("B3"),
    inspect: () => waitFrame(MAPPINGS[1].select),
  });
  assert.equal((await shapes()).length, 0);
  assert.deepEqual(
    (await wb.smoke(() => globalThis.__smoke.shapes("NK_TheoNgay"))).map((shape) => shape.name),
    ["Logo"],
  );
  assert.equal((await writesSince(start)).length, 0, "viền nháy không ghi vào ô");
  assert.equal(await targetField(d, 2).inputValue(), `${SHEET}!${MAPPINGS[1].cell}`);
  d = await wb.pickVia(d, "Chọn Ô trên form dòng 2 trên sheet", [SHEET, "A1"], { confirm: "close" });
  assert.equal(await targetField(d, 2).inputValue(), `${SHEET}!${MAPPINGS[1].cell}`);
  assert.equal(await wb.smoke(() => globalThis.__smoke.selectionHandlerCount()), 0, "gỡ sự kiện chọn sau khi xong");
  pass(
    "Gán từng dòng",
    "dòng mới để trống; ⬆ → thanh chọn, địa chỉ hiện ngay khi chọn, viền nháy quanh ô → ⬇ (ô gộp → ô trên-trái); gõ tay; X giữ ô cũ",
  );

  // 4. Dòng thừa: + Thêm rồi chọn dòng → Delete; ô trùng → báo lỗi theo dòng, không ghi.
  await d.page.getByRole("button", { name: "+ Thêm" }).click();
  assert.equal(await d.page.getByRole("row").count(), MAPPINGS.length + 2);
  await d.page.getByRole("row").last().locator("td").first().click();
  await d.page.keyboard.press("Delete");
  assert.equal(await d.page.getByRole("row").count(), MAPPINGS.length + 1);
  await d.page.getByRole("button", { name: "+ Thêm" }).click();
  const extra = MAPPINGS.length + 1;
  await chooseColumn(d, extra, "C");
  await targetField(d, extra).fill(`${SHEET}!E2`);
  start = await logLength();
  await d.click("Gán");
  d = await wb.dialog();
  assert.match(await d.text(), new RegExp(`Dòng ${extra}: ô ${SHEET}!E2 đã được gán ở dòng 1`));
  assert.equal((await writesSince(start)).length, 0);
  await d.page.getByRole("row").last().locator("td").first().click();
  await d.page.getByRole("button", { name: "Xóa dòng" }).click();
  assert.equal(await d.page.getByRole("row").count(), MAPPINGS.length + 1);
  pass("Thêm / xóa dòng gán", "Delete xóa dòng đang chọn; ô trùng → báo theo dòng, không ghi");

  // 5. Gán → form nhận dữ liệu ngày đầu tiên; chỉ ghi ô đích; công thức/tiêu đề giữ nguyên.
  start = await logLength();
  await d.click("Gán");
  d = await wb.dialog();
  assert.match(await d.text(), new RegExp(`Đã gán ${MAPPINGS.length} ô; biểu mẫu đang hiển thị dữ liệu ngày đầu tiên`));
  await d.click("Đóng");
  await done();
  await assertDayOnForm("15/06/2026");
  const writes = await writesSince(start);
  for (const w of writes) {
    const range = w.range.replace(/^(.+!)([A-Z]+\d+):\2$/, "$1$2");
    assert.ok(TARGET_CELLS.has(range), `Ghi ngoài ô đích: ${range}`);
  }
  assert.equal(writes.length, MAPPINGS.length);
  assert.equal((await cell(meta.formulaCell)).formula, formSheet.formulas[meta.formulaCell]);
  assert.equal((await cell("A1")).value, "TÌNH HÌNH THI CÔNG HÀNG NGÀY");
  assert.equal((await config()).mappings.length, MAPPINGS.length);
  assert.equal((await cell(MAPPINGS[0].cell)).value, "Ngày 15 tháng 06 năm 2026");
  pass(
    "Gán",
    "form nhận ngày đầu, dòng ngày thành 'Ngày 15 tháng 06 năm 2026'; chỉ ghi ô đích; ô công thức giữ nguyên",
  );

  // 6. Chuyển ngày sau khi gán: không mở hộp thoại; biên → thông báo.
  await pressSilently("nktcDayNext");
  await assertDayOnForm("16/06/2026");
  await pressSilently("nktcDayLast");
  await assertDayOnForm("31/12/2026");
  assert.match(await pressExpectMessage("nktcDayNext"), /Đang ở ngày cuối cùng/);
  await pressSilently("nktcDayFirst");
  assert.match(await pressExpectMessage("nktcDayPrev"), /Đang ở ngày đầu tiên/);
  pass("Chuyển ngày", "16/06 → 31/12 → (báo ngày cuối) → 15/06 → (báo ngày đầu)");

  // 7. Đọc dữ liệu mới nhất (R2.7).
  await wb.smoke(() => globalThis.__smoke.setValue("NK_TheoNgay", "C3", "Mưa to (sửa sau khi xuất)"));
  await pressSilently("nktcDayNext");
  assert.equal((await cell(MAPPINGS[2].cell)).value, "Mưa to (sửa sau khi xuất)");
  pass("Đọc lại nguồn", "sửa NK_TheoNgay!C3 → form cập nhật");

  // 7b. Cột nguồn theo tên tiêu đề + vùng nguồn tự mở rộng (R2.17): chèn cột giữa C (cột Thời tiết dịch sang D),
  // thêm 1 ngày ở cuối → Chuyển ngày vẫn ghi đúng từng cột, không phải chọn lại vùng hay gán lại.
  const setDaily = (address, value) =>
    wb.smoke(([a, v]) => globalThis.__smoke.setValue("NK_TheoNgay", a, v), [address, value]);
  await wb.smoke(() => globalThis.__smoke.insertColumn("NK_TheoNgay", 2));
  await setDaily("C1", "Ghi chú chèn thêm");
  const newRow = daily.values.length + 1;
  const NEW_DAY = 46388; // 01/01/2027
  const newValues = ["Ngày mới: nội dung", "Nắng", "Ban CH mới", 12, 3, "NT mới", "SC mới", "KN mới"];
  await setDaily(`A${newRow}`, NEW_DAY);
  await setDaily(`C${newRow}`, "không gán");
  for (let i = 0; i < newValues.length; i++)
    await setDaily(`${COLUMN_LETTERS[i < 1 ? 1 : i + 2]}${newRow}`, newValues[i]);
  start = await logLength();
  await pressSilently("nktcDayNext");
  await assertDayOnForm("17/06/2026");
  assert.equal((await writesSince(start)).length, MAPPINGS.length, "chỉ ghi các ô đích");
  const remapped = await config();
  assert.equal(remapped.headers[2], "Ghi chú chèn thêm");
  assert.deepEqual(
    remapped.mappings.map((m) => m.column),
    MAPPINGS.map((m) => COLUMN_LETTERS.indexOf(m.column) + (COLUMN_LETTERS.indexOf(m.column) >= 2 ? 1 : 0)),
  );
  await pressSilently("nktcDayLast");
  assert.equal((await cell(MAPPINGS[0].cell)).value, "Ngày 01 tháng 01 năm 2027");
  assert.equal((await cell(MAPPINGS[1].cell)).value, newValues[0]);
  assert.equal((await cell(MAPPINGS[2].cell)).value, "Nắng");
  assert.equal((await cell(MAPPINGS[8].cell)).value, "KN mới");
  pass("Chèn cột / thêm ngày", "cột khớp theo tên tiêu đề, vùng nguồn tự mở rộng tới ngày mới; không gán lại");

  // 7c. Đổi tên tiêu đề cột đã gán → Chuyển ngày chặn, báo rõ cột + ô đích, không ghi ô nào; hộp thoại Xuất vào
  // form báo dòng cần chọn lại cột. Khôi phục tên (khác hoa/thường, khoảng trắng) → chạy lại bình thường.
  await setDaily("D1", "Thời tiết (sáng)");
  start = await logLength();
  text = await pressExpectMessage("nktcDayPrev");
  assert.match(text, /Không tìm thấy cột "Thời tiết" \(gán vào Mẫu NKTC!C6\)/);
  assert.match(text, /Biểu mẫu chưa được ghi/);
  assert.equal((await writesSince(start)).length, 0);
  done = await wb.press("nktcExportForm");
  d = await wb.dialog();
  assert.match(await d.text(), /Dòng 3: cột "Thời tiết" không còn trong vùng nguồn/);
  assert.equal(await sourceField(d).inputValue(), `NK_TheoNgay!A1:J${newRow}`);
  await d.click("Hủy");
  await done();
  await setDaily("D1", "  THỜI TIẾT ");
  await pressSilently("nktcDayPrev");
  await assertDayOnForm("31/12/2026");
  pass(
    "Đổi tên tiêu đề cột",
    "Chuyển ngày chặn + báo cột/ô; hộp thoại báo dòng cần chọn lại; so tên không phân biệt hoa/thường",
  );

  // 8. Mở lại Xuất vào form: hiện cấu hình đã gán; sửa (bỏ 1 dòng) → Gán lại ghi ngày đầu.
  done = await wb.press("nktcExportForm");
  d = await wb.dialog();
  assert.equal(await d.page.getByRole("row").count(), MAPPINGS.length + 1);
  assert.equal(await targetField(d, 9).inputValue(), `${SHEET}!${MAPPINGS[8].cell}`);
  await d.page.getByRole("row").last().locator("td").first().click();
  await d.page.getByRole("button", { name: "Xóa dòng" }).click();
  await d.click("Gán");
  // Bộ ô đích thay đổi + ô đang có dữ liệu → xác nhận ghi đè (R2.5).
  d = await wb.dialog();
  assert.match(await d.text(), /sẽ bị ghi đè/);
  await d.click("Ghi đè");
  d = await wb.dialog();
  await d.click("Đóng");
  await done();
  assert.equal((await config()).mappings.length, MAPPINGS.length - 1);
  const firstRow = sourceRowForDate("15/06/2026");
  assert.ok(firstRow);
  await assertDayOnForm("15/06/2026", MAPPINGS.slice(0, -1));
  pass("Sửa cấu hình đã gán", "mở lại hiện 9 dòng; bỏ 1 dòng → Gán → xác nhận ghi đè → ngày đầu");

  // 9. Mở lại workbook: cấu hình còn; chuyển ngày không hỏi lại.
  const settings = await wb.smoke(() => globalThis.__smoke.settings());
  const sheetsNow = await wb.smoke(() => globalThis.__smoke.exportSheets());
  errors.push(...wb.pageErrors);
  await wb.context.close();
  wb = await openWorkbook(
    { sheets: sheetsNow, selection: { sheet: SHEET, address: "A1" }, apiMinor: 13, settings },
    { browser },
  );
  await pressSilently("nktcDayNext");
  assert.equal((await cell(MAPPINGS[0].cell)).value, "Ngày 16 tháng 06 năm 2026");
  pass("Mở lại workbook", "giữ gán cột và xác nhận → Ngày sau ghi 16/06 không hỏi lại");

  // 9b. Excel cũ (DialogApi 1.1): thanh chọn không cập nhật tại chỗ → bấm ⬇ lấy ô đang chọn trên sheet;
  // nút Ribbon chuyển ngày khi cần xác nhận ghi đè → hỏi rồi ghi.
  const settingsOld = await wb.smoke(() => globalThis.__smoke.settings());
  delete settingsOld[SETTINGS_KEY].confirmedKey;
  const sheetsOld = await wb.smoke(() => globalThis.__smoke.exportSheets());
  errors.push(...wb.pageErrors);
  await wb.context.close();
  wb = await openWorkbook(
    {
      sheets: sheetsOld,
      selection: { sheet: SHEET, address: "A1" },
      apiMinor: 13,
      dialogApiMinor: 1,
      settings: settingsOld,
    },
    { browser },
  );
  done = await wb.press("nktcExportForm");
  d = await wb.dialog();
  d = await wb.pickVia(d, "Chọn Ô trên form dòng 1 trên sheet", [SHEET, "B3"], {
    initial: `${SHEET}!${MAPPINGS[0].cell}`,
    inspect: async (bar) => {
      assert.equal(await wb.smoke(() => globalThis.__smoke.selectionHandlerCount()), 0, "không nghe sự kiện chọn");
      assert.match(await bar.text(), /rồi bấm ⬇/);
    },
  });
  assert.equal(await targetField(d, 1).inputValue(), `${SHEET}!B3`);
  // Đóng hộp thoại (X) → cấu hình đã gán giữ nguyên.
  await wb.smoke(() => globalThis.__smoke.dismiss());
  await d.page.close();
  await done();
  done = await wb.press("nktcDayNext");
  d = await wb.dialog();
  assert.match(await d.text(), /sẽ bị ghi đè/);
  await d.click("Ghi đè");
  await done();
  assert.equal((await cell(MAPPINGS[0].cell)).value, "Ngày 17 tháng 06 năm 2026");
  pass("Excel cũ (DialogApi 1.1)", "thanh chọn: ⬇ lấy ô đang chọn trên sheet; chuyển ngày cần xác nhận → hỏi rồi ghi");

  // 10. Vùng nguồn có dòng tiêu đề bị gộp (merge) → báo không chọn được, phải gỡ merge.
  const merged = await openWorkbook(
    {
      sheets: [
        {
          name: "NK",
          values: [
            ["Ngày", "", "Thời tiết"],
            [46188, "a", "Nắng"],
          ],
          merged: ["A1:B1"],
        },
        { name: "F", values: [[""]] },
      ],
      selection: { sheet: "NK", address: "A2" },
      apiMinor: 13,
    },
    { browser },
  );
  done = await merged.press("nktcExportForm");
  d = await merged.dialog();
  assert.match(await d.text(), /ô bị gộp \(merge\).*gỡ merge/s);
  assert.ok(await d.page.getByRole("button", { name: "Gán", exact: true }).isDisabled());
  await d.click("Hủy");
  await done();
  errors.push(...merged.pageErrors);
  await merged.context.close();
  pass("Tiêu đề có ô gộp", "báo phải gỡ merge rồi chọn lại; nút Gán bị khóa");

  // 11. Excel cũ (ExcelApi 1.1): Chèn form mẫu → hướng dẫn tải file; tiêu đề trống (merge) vẫn bị chặn.
  const old = await openWorkbook(
    {
      sheets: [
        {
          name: "NK",
          values: [
            ["Ngày", "", "Thời tiết"],
            [46188, "a", "Nắng"],
          ],
        },
      ],
      selection: { sheet: "NK", address: "A2" },
      apiMinor: 1,
    },
    { browser },
  );
  done = await old.press("nktcInsertTemplate");
  d = await old.dialog();
  text = await d.text();
  assert.match(text, /chưa hỗ trợ chèn sheet/);
  assert.equal(
    await d.page.getByRole("link", { name: /Mau_NKTC\.xlsx/ }).getAttribute("href"),
    "templates/Mau_NKTC.xlsx",
  );
  await d.click("Đóng");
  await done();
  done = await old.press("nktcExportForm");
  d = await old.dialog();
  assert.match(await d.text(), /ô tiêu đề trống \(cột thứ 2\)/);
  await d.click("Hủy");
  await done();
  errors.push(...old.pageErrors);
  await old.context.close();
  pass("Excel cũ (ExcelApi 1.1)", "Chèn form mẫu → link tải Mau_NKTC.xlsx; tiêu đề trống → chặn");

  // 12. Sheet biểu mẫu bị đổi tên → báo lỗi, không ghi ô nào (R2.9).
  await wb.smoke((s) => globalThis.__smoke.renameSheet(s, `${s} cũ`), SHEET);
  start = await logLength();
  text = await pressExpectMessage("nktcDayNext");
  assert.ok(text.includes(`Không tìm thấy sheet: "${SHEET}"`), text);
  assert.equal((await writesSince(start)).length, 0);
  pass("Sheet biểu mẫu bị đổi tên", "báo lỗi, không ghi ô nào");

  errors.push(...wb.pageErrors);
  await browser.close();
  console.log(`Smoke F2 — ${steps.length} bước:\n${steps.join("\n")}`);
  console.log("PAGE ERRORS:", errors.length ? errors : "none");
  if (errors.length) process.exitCode = 1;
}

fs.mkdirSync("smoke-output", { recursive: true });
main().catch((error) => {
  console.log(steps.join("\n"));
  console.error("SMOKE F2 FAIL:", error.message);
  process.exit(1);
});
