/*
 * Smoke test F2 (nhóm Form mẫu + Chuyển ngày trên Ribbon) trên bundle production với Excel giả lập:
 * Chèn form mẫu, hộp thoại chung "Xuất vào form" (vùng nguồn + dòng gán + Gán), chuyển ngày.
 * Dữ liệu: test/fixtures/demo_f2.json + demo_f2.targets.json (ô đích của từng cột).
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
  async function assertDayOnForm(label) {
    const row = sourceRowForDate(label);
    for (const m of MAPPINGS) {
      const expected = row[COLUMN_LETTERS.indexOf(m.column)];
      const actual = (await cell(m.cell))?.value ?? "";
      assert.deepEqual(actual, expected, `${SHEET}!${m.cell} (cột ${m.column} – ${m.header}) ngày ${label}`);
    }
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
  pass("Chèn form mẫu", `sheet "${SHEET}" từ templates/Mau_NKTC.xlsx; lần 2 báo đã có, không ghi đè`);

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
  pass("Mở Xuất vào form", `vùng nguồn NK_TheoNgay!${sourceAddress} (gồm cột đã thêm), cột Ngày A, 200 ngày`);

  // 3. Gán từng dòng: chọn cột (trái), chọn ô trên form rồi bấm ⬆ (phải; ô gộp → ô trên-trái); + Thêm.
  for (let i = 0; i < MAPPINGS.length; i++) {
    const m = MAPPINGS[i];
    const row = i + 1;
    if (i > 0) await d.page.getByRole("button", { name: "+ Thêm" }).click();
    await chooseColumn(d, row, m.column);
    await select(SHEET, m.select);
    await d.act((page) => page.getByRole("button", { name: `Chọn Ô trên form dòng ${row} trên sheet` }).click());
    d = await wb.dialog();
    assert.equal(await targetField(d, row).inputValue(), `${SHEET}!${m.cell}`);
  }
  await d.page.screenshot({ path: path.join("smoke-output", "f2-export-form.png"), fullPage: true });
  pass(`Gán ${MAPPINGS.length} dòng`, MAPPINGS.map((m) => `${m.column}→${m.cell}`).join(", "));

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
  let start = await logLength();
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
  pass(
    "Gán",
    `form nhận ngày 15/06/2026; ${writes.length} lệnh ghi đúng ô đích; công thức ${meta.formulaCell} giữ nguyên`,
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

  // 7. Đọc dữ liệu mới nhất (R2.7) + Đến ngày.
  await wb.smoke(() => globalThis.__smoke.setValue("NK_TheoNgay", "C3", "Mưa to (sửa sau khi xuất)"));
  await pressSilently("nktcDayNext");
  assert.equal((await cell(MAPPINGS[2].cell)).value, "Mưa to (sửa sau khi xuất)");
  done = await wb.press("nktcGoToDay");
  d = await wb.dialog();
  await d.page.getByRole("textbox", { name: "Ngày (dd/mm/yyyy)" }).fill("05/08/2026");
  await d.click("Đi tới");
  await done();
  await assertDayOnForm("05/08/2026");
  pass("Đọc lại nguồn + Đến ngày", "sửa NK_TheoNgay!C3 → form cập nhật; 05/08/2026");

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
  for (const m of MAPPINGS.slice(0, -1)) {
    assert.deepEqual((await cell(m.cell))?.value ?? "", firstRow[COLUMN_LETTERS.indexOf(m.column)], m.cell);
  }
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
  assert.equal((await cell(MAPPINGS[0].cell)).value, sourceRowForDate("16/06/2026")[0]);
  pass("Mở lại workbook", "giữ gán cột và xác nhận → Ngày sau ghi 16/06 không hỏi lại");

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
