/*
 * Smoke test F2 (xuất theo form mẫu) trên bundle production với Excel giả lập.
 * Dữ liệu: test/fixtures/demo_f2.json (sinh từ demo_f2.xlsx) + demo_f2.targets.json (ô đích của từng cột).
 * Mỗi bước có kiểm tra; sai → exit code 1.
 *
 *   node scripts/smoke/run-f2.js [fixture.json] [--api 1]
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { openTaskpane } = require("./harness");

const fixturePath =
  process.argv[2] && !process.argv[2].startsWith("--") ? process.argv[2] : "test/fixtures/demo_f2.json";
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
  let session = await openTaskpane({
    sheets: fixture.sheets,
    selection: { sheet: "NK_TheoNgay", address: sourceAddress },
    apiMinor,
  });
  let { page } = session;
  const smoke = (fn, arg) => page.evaluate(fn, arg);
  const cell = (address) => smoke(([s, a]) => globalThis.__smoke.cell(s, a), [SHEET, address]);
  const select = (address) => smoke(([s, a]) => globalThis.__smoke.select(s, a), [SHEET, address]);
  const logLength = () => smoke(() => globalThis.__smoke.log.length);
  const writesSince = (start) => smoke((s) => globalThis.__smoke.log.slice(s).filter((e) => e.op === "values"), start);
  const lastMessage = () => page.locator(".ms-MessageBar").last().innerText();
  const currentDay = () => page.getByTestId("current-day").innerText();
  const settle = () => page.waitForTimeout(150);
  const waitDay = (label) =>
    page.waitForFunction(
      (l) => globalThis.document.querySelector('[data-testid="current-day"]')?.textContent === l,
      label,
    );
  const assign = (column) => page.getByRole("button", { name: `Gán ô đang chọn cho cột ${column}` }).click();

  async function assertDayOnForm(label) {
    const row = sourceRowForDate(label);
    assert.equal((await currentDay()).trim(), label);
    for (const m of MAPPINGS) {
      const expected = row[COLUMN_LETTERS.indexOf(m.column)];
      const actual = (await cell(m.cell))?.value ?? "";
      assert.deepEqual(actual, expected, `${SHEET}!${m.cell} (cột ${m.column} – ${m.header}) ngày ${label}`);
    }
  }

  // 1. Chọn vùng nguồn.
  await page.getByRole("tab", { name: "Form mẫu" }).click();
  await page.getByRole("button", { name: "Lấy vùng đang chọn" }).click();
  await page.waitForSelector("text=Vùng nguồn:");
  const sourceText = /Vùng nguồn:[^\n]*/.exec(await page.locator("#root").innerText())?.[0] ?? "";
  assert.ok(sourceText.includes(`NK_TheoNgay!${sourceAddress} · ${MAPPINGS.length} cột · 200 ngày`), sourceText);
  pass("Chọn vùng nguồn", sourceText.trim());

  // 2. Gán cột → ô (chọn ô merge → lấy ô trên-trái, R2.2). Tên sheet có dấu và khoảng trắng.
  for (const m of MAPPINGS) {
    await select(m.select);
    await assign(m.column);
    await page.waitForSelector(`text=→ ${SHEET}!${m.cell}`);
  }
  pass(`Gán ${MAPPINGS.length} cột vào "${SHEET}"`, MAPPINGS.map((m) => `${m.column}→${m.cell}`).join(", "));

  // 3. Ô đích đã gán cho cột khác → cảnh báo, không gán (R2.12).
  await select(MAPPINGS[4].select);
  await assign(MAPPINGS[2].column);
  await settle();
  assert.match(await lastMessage(), /đã được gán cho cột/);
  pass("Chặn gán trùng ô đích", (await lastMessage()).trim());

  // 4. Xuất ngày đầu: chỉ ghi vào ô đích; công thức và tiêu đề giữ nguyên (R2.3, R2.6).
  let start = await logLength();
  await page.getByRole("button", { name: "Xuất lên form" }).click();
  await page.getByTestId("current-day").waitFor();
  await assertDayOnForm("15/06/2026");
  const writes = await writesSince(start);
  // Nhãn ghi của Excel giả lập có dạng "Sheet!B4:B4" → chuẩn hóa thành "Sheet!B4".
  const written = writes.map((w) => w.range.replace(/^(.+!)([A-Z]+\d+):\2$/, "$1$2"));
  for (const range of written) assert.ok(TARGET_CELLS.has(range), `Ghi ngoài ô đích: ${range}`);
  assert.equal(writes.length, MAPPINGS.length);
  assert.equal((await cell(meta.formulaCell)).formula, formSheet.formulas[meta.formulaCell]);
  assert.equal((await cell("A1")).value, "TÌNH HÌNH THI CÔNG HÀNG NGÀY");
  pass(
    "Xuất ngày đầu 15/06/2026",
    `${writes.length} lệnh ghi, đúng ${MAPPINGS.length} ô đích; công thức ${meta.formulaCell} và tiêu đề A1 giữ nguyên`,
  );

  // 5. Văn bản bắt đầu bằng "- " được ghi nguyên văn (R2.11).
  const content = (await cell(MAPPINGS[1].cell)).value;
  assert.ok(content.startsWith("- HM - CÔNG TÁC CHUẨN BỊ"));
  pass("Nội dung nhiều dòng ghi nguyên văn", JSON.stringify(content.split("\n")[0]));

  // 6. Điều hướng: sau / cuối / đầu.
  await page.getByRole("button", { name: "Ngày sau" }).click();
  await waitDay("16/06/2026");
  await assertDayOnForm("16/06/2026");
  await page.getByRole("button", { name: "Ngày cuối" }).click();
  await waitDay("31/12/2026");
  await assertDayOnForm("31/12/2026");
  assert.ok(await page.getByRole("button", { name: "Ngày sau" }).isDisabled());
  await page.getByRole("button", { name: "Ngày đầu" }).click();
  await waitDay("15/06/2026");
  assert.ok(await page.getByRole("button", { name: "Ngày trước" }).isDisabled());
  pass("Điều hướng Sau / Cuối / Đầu", "16/06 → 31/12 → 15/06; nút bị khóa ở hai đầu");

  // 7. Đọc dữ liệu mới nhất mỗi lần chuyển ngày (R2.7).
  await smoke(() => globalThis.__smoke.setValue("NK_TheoNgay", "C3", "Mưa to (sửa sau khi xuất)"));
  await page.getByRole("button", { name: "Ngày sau" }).click();
  await waitDay("16/06/2026");
  assert.equal((await cell(MAPPINGS[2].cell)).value, "Mưa to (sửa sau khi xuất)");
  pass("Đọc lại vùng nguồn mới nhất", `sửa NK_TheoNgay!C3 → ${SHEET}!${MAPPINGS[2].cell} cập nhật`);

  // 8. Nhảy tới ngày; ngày sai; ngày ngoài phạm vi.
  const jump = async (text) => {
    await page.getByRole("textbox", { name: "Nhảy tới ngày" }).fill(text);
    await page.getByRole("button", { name: "Đi tới" }).click();
    await settle();
  };
  await jump("05/08/2026");
  await waitDay("05/08/2026");
  await assertDayOnForm("05/08/2026");
  assert.equal((await cell(MAPPINGS[1].cell)).value.split("\n").length, 13);
  await jump("31/02/2026");
  assert.match(await lastMessage(), /không hợp lệ/);
  await jump("01/01/2027");
  assert.match(await lastMessage(), /không có trong vùng nguồn/);
  pass("Nhảy tới ngày", "05/08/2026 (13 dòng nội dung, có sự cố + kiến nghị); báo lỗi 31/02/2026 và 01/01/2027");

  // 9. Mở lại task pane: cấu hình được khôi phục (R2.8); ô đích đã có dữ liệu → xin xác nhận (R2.5).
  const settings = await smoke(() => globalThis.__smoke.settings());
  const sheetsNow = await smoke(() => globalThis.__smoke.exportSheets());
  const browser = session.browser;
  await page.close();
  session = await openTaskpane(
    { sheets: sheetsNow, selection: { sheet: SHEET, address: "A1" }, apiMinor, settings },
    { browser },
  );
  page = session.page;
  await page.getByRole("tab", { name: "Form mẫu" }).click();
  await page.waitForSelector(`text=→ ${SHEET}!${MAPPINGS[MAPPINGS.length - 1].cell}`);
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.waitForSelector("text=sẽ bị ghi đè");
  start = await logLength();
  await page.getByRole("button", { name: "Hủy" }).click();
  await settle();
  assert.equal((await writesSince(start)).length, 0);
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.getByRole("button", { name: "Ghi đè" }).click();
  await page.getByTestId("current-day").waitFor();
  await assertDayOnForm("05/08/2026");
  pass(
    "Khôi phục cấu hình khi mở lại",
    `${MAPPINGS.length} gán cột + ngày 05/08/2026; xác nhận ghi đè: Hủy không ghi, Ghi đè thì ghi`,
  );

  // 10. Ô đích chứa công thức → cảnh báo CÔNG THỨC; Hủy giữ nguyên công thức (R2.4).
  await select(meta.formulaCell);
  await assign("A");
  await page.waitForSelector(`text=→ ${SHEET}!${meta.formulaCell}`);
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.waitForSelector("text=CÔNG THỨC");
  await page.getByRole("button", { name: "Hủy" }).click();
  await settle();
  assert.equal((await cell(meta.formulaCell)).formula, formSheet.formulas[meta.formulaCell]);
  await page.getByRole("button", { name: `Bỏ gán ô ${SHEET}!${meta.formulaCell}` }).click();
  await settle();
  pass("Cảnh báo ô đích chứa công thức", `Hủy → công thức ${meta.formulaCell} (ô tích thời tiết) giữ nguyên; bỏ gán`);

  // 11. Sheet biểu mẫu bị đổi tên → báo lỗi, không ghi ô nào (R2.9).
  await smoke((s) => globalThis.__smoke.renameSheet(s, `${s} cũ`), SHEET);
  start = await logLength();
  await page.getByRole("button", { name: "Ngày sau" }).click();
  await settle();
  assert.ok((await lastMessage()).includes(`Không tìm thấy sheet: "${SHEET}"`), await lastMessage());
  assert.equal((await writesSince(start)).length, 0);
  pass("Sheet biểu mẫu bị đổi tên", (await lastMessage()).trim());

  await page.screenshot({ path: path.join("smoke-output", "f2-panel.png"), fullPage: true });
  const errors = session.pageErrors;
  await session.browser.close();
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
