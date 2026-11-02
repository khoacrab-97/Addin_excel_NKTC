/*
 * Smoke test F2 (nhóm Form mẫu + Chuyển ngày trên Ribbon) trên bundle production với Excel giả lập.
 * Dữ liệu: test/fixtures/demo_f2.json (sinh từ demo_f2.xlsx) + demo_f2.targets.json (ô đích của từng cột).
 * Mỗi bước có kiểm tra; sai → exit code 1.
 *
 *   node scripts/smoke/run-f2.js [tên file trong test/fixtures] [--api 1]
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
const apiIndex = process.argv.indexOf("--api");
const apiMinor = apiIndex === -1 ? 1 : Number(process.argv[apiIndex + 1]);
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
  let wb = await openWorkbook(
    { sheets: fixture.sheets, selection: { sheet: "NK_TheoNgay", address: "B7" }, apiMinor },
    { browser },
  );
  const errors = [];
  const cell = (address) => wb.smoke(([s, a]) => globalThis.__smoke.cell(s, a), [SHEET, address]);
  const select = (sheet, address) => wb.smoke(([s, a]) => globalThis.__smoke.select(s, a), [sheet, address]);
  const logLength = () => wb.smoke(() => globalThis.__smoke.log.length);
  const writesSince = (start) =>
    wb.smoke((s) => globalThis.__smoke.log.slice(s).filter((e) => e.op === "values"), start);
  const config = async () => (await wb.smoke(() => globalThis.__smoke.settings()))[SETTINGS_KEY];

  /** Bấm nút; nếu có hộp thoại thông báo thì trả về nội dung rồi đóng. */
  async function pressExpectMessage(name) {
    const done = await wb.press(name);
    const d = await wb.dialog();
    const text = await d.text();
    await d.click("Đóng");
    await done();
    return text;
  }
  /** Bấm nút chuyển ngày, không có hộp thoại nào. */
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
  async function assign(m, column = m.column) {
    await select(SHEET, m.select);
    const done = await wb.press("nktcAssignCell");
    const d = await wb.dialog();
    assert.ok((await d.text()).includes(`Ô trên biểu mẫu: ${SHEET}!${m.cell}`), await d.text());
    await d.choose("Cột ghi vào ô này", `${column} –`);
    await d.click("Gán");
    await done();
  }

  // 0. Chưa có cấu hình → nút Gán cột / chuyển ngày hướng dẫn đặt vùng nguồn.
  let text = await pressExpectMessage("nktcDayNext");
  assert.match(text, /Chưa có cấu hình form mẫu/);
  pass("Chưa cấu hình", "hướng dẫn bấm Vùng nguồn trước");

  // 1. Vùng nguồn: đặt con trỏ 1 ô trong NK_TheoNgay → lấy cả bảng, nhận diện cột Ngày.
  text = await pressExpectMessage("nktcSetSource");
  assert.ok(text.includes(`Vùng nguồn: NK_TheoNgay!${sourceAddress}.`), text);
  assert.ok(text.includes("Cột Ngày: A – Ngày") && text.includes("200 ngày hợp lệ"), text);
  pass("Vùng nguồn", `NK_TheoNgay!${sourceAddress}, cột Ngày A, 200 ngày`);

  // 2. Gán cột → ô (chọn ô merge → lấy ô trên-trái, R2.2). Tên sheet có dấu và khoảng trắng.
  for (const m of MAPPINGS) await assign(m);
  let saved = await config();
  assert.equal(saved.mappings.length, MAPPINGS.length);
  pass(`Gán ${MAPPINGS.length} cột vào "${SHEET}"`, MAPPINGS.map((m) => `${m.column}→${m.cell}`).join(", "));

  // 3. Gán lại ô đã gán: hộp thoại hiện cột hiện tại, đổi cột thì thay (mỗi ô một cột, R2.12).
  await select(SHEET, MAPPINGS[4].select);
  let done = await wb.press("nktcAssignCell");
  let d = await wb.dialog();
  assert.match(await d.text(), /Ô này đang được gán/);
  assert.ok(
    (await d.page.getByRole("combobox", { name: "Cột ghi vào ô này" }).innerText()).startsWith(
      `${MAPPINGS[4].column} –`,
    ),
  );
  await d.choose("Cột ghi vào ô này", `${MAPPINGS[2].column} –`);
  await d.click("Gán");
  await done();
  saved = await config();
  assert.equal(saved.mappings.length, MAPPINGS.length);
  assert.equal(
    saved.mappings.find((x) => x.target.address === MAPPINGS[4].cell).column,
    COLUMN_LETTERS.indexOf(MAPPINGS[2].column),
  );
  await assign(MAPPINGS[4]);
  pass("Gán lại ô đã gán", "hiện cột đang gán, đổi cột thay cho cũ, không tạo gán trùng");

  // 4. Ô đích đã có dữ liệu → hỏi xác nhận (R2.5): Hủy không ghi; Ghi đè thì ghi đúng các ô đích (R2.3, R2.6).
  await wb.smoke(([s, a]) => globalThis.__smoke.setValue(s, a, "dữ liệu cũ"), [SHEET, MAPPINGS[2].cell]);
  let start = await logLength();
  done = await wb.press("nktcDayFirst");
  d = await wb.dialog();
  text = await d.text();
  assert.ok(text.includes(`${SHEET}!${MAPPINGS[2].cell} — có dữ liệu`) && text.includes("15/06/2026"), text);
  await d.page.screenshot({ path: path.join("smoke-output", "f2-confirm.png") });
  await d.click("Hủy");
  await done();
  assert.equal((await writesSince(start)).length, 0);
  assert.equal((await cell(MAPPINGS[2].cell)).value, "dữ liệu cũ");
  done = await wb.press("nktcDayFirst");
  d = await wb.dialog();
  await d.click("Ghi đè");
  await done();
  await assertDayOnForm("15/06/2026");
  const writes = await writesSince(start);
  const written = writes.map((w) => w.range.replace(/^(.+!)([A-Z]+\d+):\2$/, "$1$2"));
  for (const range of written) assert.ok(TARGET_CELLS.has(range), `Ghi ngoài ô đích: ${range}`);
  assert.equal(writes.length, MAPPINGS.length);
  assert.equal((await cell(meta.formulaCell)).formula, formSheet.formulas[meta.formulaCell]);
  assert.equal((await cell("A1")).value, "TÌNH HÌNH THI CÔNG HÀNG NGÀY");
  assert.equal((await config()).currentDay, 46188);
  pass(
    "Ngày đầu + xác nhận ghi đè",
    `Hủy → không ghi; Ghi đè → ${writes.length} lệnh ghi đúng ô đích; công thức ${meta.formulaCell}, tiêu đề A1 giữ nguyên`,
  );

  // 5. Văn bản bắt đầu bằng "- " được ghi nguyên văn (R2.11).
  const content = (await cell(MAPPINGS[1].cell)).value;
  assert.ok(content.startsWith("- HM - CÔNG TÁC CHUẨN BỊ"));
  pass("Nội dung nhiều dòng ghi nguyên văn", JSON.stringify(content.split("\n")[0]));

  // 6. Chuyển ngày: không hỏi lại (đã xác nhận), không mở hộp thoại; ở hai đầu → thông báo.
  await pressSilently("nktcDayNext");
  await assertDayOnForm("16/06/2026");
  await pressSilently("nktcDayLast");
  await assertDayOnForm("31/12/2026");
  assert.match(await pressExpectMessage("nktcDayNext"), /Đang ở ngày cuối cùng \(31\/12\/2026\)/);
  await pressSilently("nktcDayFirst");
  await assertDayOnForm("15/06/2026");
  assert.match(await pressExpectMessage("nktcDayPrev"), /Đang ở ngày đầu tiên/);
  pass("Ngày sau / cuối / đầu / trước", "16/06 → 31/12 → (báo ngày cuối) → 15/06 → (báo ngày đầu); không hỏi lại");

  // 7. Đọc dữ liệu mới nhất mỗi lần chuyển ngày (R2.7).
  await wb.smoke(() => globalThis.__smoke.setValue("NK_TheoNgay", "C3", "Mưa to (sửa sau khi xuất)"));
  await pressSilently("nktcDayNext");
  assert.equal((await cell(MAPPINGS[2].cell)).value, "Mưa to (sửa sau khi xuất)");
  pass("Đọc lại vùng nguồn mới nhất", `sửa NK_TheoNgay!C3 → ${SHEET}!${MAPPINGS[2].cell} cập nhật`);

  // 8. Đến ngày: hợp lệ; sai định dạng (nút khóa); ngoài phạm vi (báo lỗi).
  done = await wb.press("nktcGoToDay");
  d = await wb.dialog();
  assert.match(await d.text(), /từ 15\/06\/2026 đến 31\/12\/2026/);
  const box = d.page.getByRole("textbox", { name: "Ngày (dd/mm/yyyy)" });
  await box.fill("31/02/2026");
  assert.ok(await d.page.getByRole("button", { name: "Đi tới" }).isDisabled());
  await box.fill("05/08/2026");
  await d.page.screenshot({ path: path.join("smoke-output", "f2-goto.png") });
  await d.click("Đi tới");
  await done();
  await assertDayOnForm("05/08/2026");
  assert.equal((await cell(MAPPINGS[1].cell)).value.split("\n").length, 13);
  done = await wb.press("nktcGoToDay");
  d = await wb.dialog();
  await d.page.getByRole("textbox", { name: "Ngày (dd/mm/yyyy)" }).fill("01/01/2027");
  await d.click("Đi tới");
  d = await wb.dialog();
  assert.match(await d.text(), /không có trong vùng nguồn/);
  await d.click("Đóng");
  await done();
  pass("Đến ngày", "05/08/2026 (13 dòng nội dung); 31/02/2026 khóa nút; 01/01/2027 báo không có");

  // 9. Mở lại workbook: cấu hình + xác nhận được khôi phục (R2.8) → chuyển ngày không hỏi lại.
  const settings = await wb.smoke(() => globalThis.__smoke.settings());
  const sheetsNow = await wb.smoke(() => globalThis.__smoke.exportSheets());
  errors.push(...wb.pageErrors);
  await wb.context.close();
  wb = await openWorkbook(
    { sheets: sheetsNow, selection: { sheet: SHEET, address: "A1" }, apiMinor, settings },
    { browser },
  );
  await pressSilently("nktcDayPrev");
  await assertDayOnForm("04/08/2026");
  pass("Mở lại workbook", "khôi phục gán cột, ngày 05/08 và xác nhận → Ngày trước ghi 04/08 không hỏi lại");

  // 10. Thêm ô có công thức → bộ ô đích đổi → hỏi lại, nêu rõ CÔNG THỨC; Hủy giữ nguyên (R2.4).
  await select(SHEET, meta.formulaCell);
  done = await wb.press("nktcAssignCell");
  d = await wb.dialog();
  await d.choose("Cột ghi vào ô này", "A –");
  await d.click("Gán");
  await done();
  start = await logLength();
  done = await wb.press("nktcDayNext");
  d = await wb.dialog();
  assert.ok((await d.text()).includes(`${SHEET}!${meta.formulaCell} — có công thức`));
  await d.click("Hủy");
  await done();
  assert.equal((await writesSince(start)).length, 0);
  assert.equal((await cell(meta.formulaCell)).formula, formSheet.formulas[meta.formulaCell]);
  pass("Cảnh báo ô đích chứa công thức", `Hủy → công thức ${meta.formulaCell} giữ nguyên`);

  // 11. Danh sách gán: bỏ gán ô công thức. Gán cột đã thay đổi nên xác nhận cũ hết hiệu lực →
  //     lần ghi kế tiếp hỏi lại (ô đích đang có dữ liệu), Ghi đè → các lần sau không hỏi.
  done = await wb.press("nktcMappingList");
  d = await wb.dialog();
  text = await d.text();
  assert.ok(text.includes(`${SHEET}!${meta.formulaCell}`) && text.includes("A – Ngày"), text);
  await d.page.screenshot({ path: path.join("smoke-output", "f2-settings.png") });
  await d.page.getByRole("checkbox", { name: `Bỏ gán ${SHEET}!${meta.formulaCell}` }).check({ force: true });
  await d.click("Lưu");
  await done();
  assert.equal((await config()).mappings.length, MAPPINGS.length);
  done = await wb.press("nktcDayNext");
  d = await wb.dialog();
  assert.ok(!(await d.text()).includes(meta.formulaCell));
  await d.click("Ghi đè");
  await done();
  await assertDayOnForm("05/08/2026");
  await pressSilently("nktcDayNext");
  await assertDayOnForm("06/08/2026");
  pass("Danh sách gán", `bỏ gán ${meta.formulaCell}; gán cột đổi → hỏi xác nhận 1 lần, sau đó ghi thẳng`);

  // 12. Bấm nút khác khi đang mở hộp thoại → bỏ qua (một lệnh tại một thời điểm).
  done = await wb.press("nktcHelp");
  d = await wb.dialog();
  start = await logLength();
  const second = await wb.press("nktcDayNext");
  await second();
  assert.equal((await writesSince(start)).length, 0);
  await d.click("Đóng");
  await done();
  pass("Một lệnh tại một thời điểm", "bấm Ngày sau khi đang mở Hướng dẫn → bỏ qua");

  // 13. Sheet biểu mẫu bị đổi tên → báo lỗi, không ghi ô nào (R2.9).
  await wb.smoke((s) => globalThis.__smoke.renameSheet(s, `${s} cũ`), SHEET);
  start = await logLength();
  text = await pressExpectMessage("nktcDayNext");
  assert.ok(text.includes(`Không tìm thấy sheet: "${SHEET}"`), text);
  assert.equal((await writesSince(start)).length, 0);
  pass(
    "Sheet biểu mẫu bị đổi tên",
    text.split("\n").find((line) => line.includes("Không tìm thấy")),
  );

  errors.push(...wb.pageErrors);
  await browser.close();
  console.log(`Smoke F2 (ExcelApi 1.${apiMinor}) — ${steps.length} bước:\n${steps.join("\n")}`);
  console.log("PAGE ERRORS:", errors.length ? errors : "none");
  if (errors.length) process.exitCode = 1;
}

fs.mkdirSync("smoke-output", { recursive: true });
main().catch((error) => {
  console.log(steps.join("\n"));
  console.error("SMOKE F2 FAIL:", error.message);
  process.exit(1);
});
