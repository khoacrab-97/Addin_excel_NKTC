/*
 * Smoke test F2 (xuất theo form mẫu) trên bundle production với Excel giả lập.
 * Dữ liệu: test/fixtures/demo_f2.json (sinh từ demo_f2.xlsx). Mỗi bước có kiểm tra; sai → exit code 1.
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
const daily = fixture.sheets.find((sheet) => sheet.name === "NK_TheoNgay");

/** Cột nguồn → vùng chọn trên Form (vùng merge, như Excel trả về khi chọn ô merge) → ô ghi kỳ vọng. */
const MAPPINGS = [
  { column: "A", select: "B4:C4", cell: "B4" },
  { column: "B", select: "A9:F18", cell: "A9" },
  { column: "C", select: "E4:F4", cell: "E4" },
  { column: "D", select: "B5:F5", cell: "B5" },
  { column: "E", select: "B6", cell: "B6" },
  { column: "F", select: "E6", cell: "E6" },
  { column: "G", select: "A21:F24", cell: "A21" },
];
const TARGET_CELLS = new Set(MAPPINGS.map((m) => `Form!${m.cell}`));

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
    selection: { sheet: "NK_TheoNgay", address: "A1:G201" },
    apiMinor,
  });
  let { page } = session;
  const smoke = (fn, ...args) => page.evaluate(fn, ...args);
  const cell = (address) => smoke((a) => globalThis.__smoke.cell("Form", a), address);
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

  async function assertDayOnForm(label) {
    const row = sourceRowForDate(label);
    assert.equal((await currentDay()).trim(), label);
    for (const m of MAPPINGS) {
      const expected = row["ABCDEFG".indexOf(m.column)];
      const actual = (await cell(m.cell))?.value ?? "";
      assert.deepEqual(actual, expected, `Form!${m.cell} (cột ${m.column}) ngày ${label}`);
    }
  }

  // 1. Chọn vùng nguồn.
  await page.getByRole("tab", { name: "Form mẫu" }).click();
  await page.getByRole("button", { name: "Lấy vùng đang chọn" }).click();
  await page.waitForSelector("text=Vùng nguồn:");
  const sourceText = /Vùng nguồn:[^\n]*/.exec(await page.locator("#root").innerText())?.[0] ?? "";
  assert.match(sourceText, /NK_TheoNgay!A1:G201 · 7 cột · 200 ngày/);
  pass("Chọn vùng nguồn", sourceText.trim());

  // 2. Gán cột → ô (chọn ô merge → lấy ô trên-trái, R2.2).
  for (const m of MAPPINGS) {
    await smoke(([s, a]) => globalThis.__smoke.select(s, a), ["Form", m.select]);
    await page.getByRole("button", { name: `Gán ô đang chọn cho cột ${m.column}` }).click();
    await page.waitForSelector(`text=→ Form!${m.cell}`);
  }
  pass("Gán 7 cột vào ô form", MAPPINGS.map((m) => `${m.column}→${m.cell}`).join(", "));

  // 3. Ô đích đã gán cho cột khác → cảnh báo, không gán (R2.1).
  await smoke(() => globalThis.__smoke.select("Form", "B6"));
  await page.getByRole("button", { name: "Gán ô đang chọn cho cột C" }).click();
  await settle();
  assert.match(await lastMessage(), /đã được gán cho cột/);
  pass("Chặn gán trùng ô đích", (await lastMessage()).trim());

  // 4. Xuất ngày đầu: chỉ ghi vào ô đích, giữ công thức ở A26 (R2.3, R2.6).
  let start = await logLength();
  await page.getByRole("button", { name: "Xuất lên form" }).click();
  await page.getByTestId("current-day").waitFor();
  await assertDayOnForm("15/06/2026");
  const writes = await writesSince(start);
  // Nhãn ghi của Excel giả lập có dạng "Form!B4:B4" → chuẩn hóa thành "Form!B4".
  const written = writes.map((w) => w.range.replace(/^(.+!)([A-Z]+\d+):\2$/, "$1$2"));
  for (const range of written) assert.ok(TARGET_CELLS.has(range), `Ghi ngoài ô đích: ${range}`);
  assert.equal(writes.length, MAPPINGS.length);
  assert.equal((await cell("A26")).formula, fixture.sheets.find((s) => s.name === "Form").formulas.A26);
  assert.equal((await cell("A4")).value, "Ngày:");
  pass(
    "Xuất ngày đầu 15/06/2026",
    `${writes.length} lệnh ghi, đúng ${MAPPINGS.length} ô đích; công thức A26 và nhãn A4 giữ nguyên`,
  );

  // 5. Văn bản bắt đầu bằng "- " được ghi nguyên văn (R2.11).
  const content = (await cell("A9")).value;
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
  assert.equal((await cell("E4")).value, "Mưa to (sửa sau khi xuất)");
  pass("Đọc lại vùng nguồn mới nhất", "sửa NK_TheoNgay!C3 → Form!E4 cập nhật");

  // 8. Nhảy tới ngày; ngày sai; ngày ngoài phạm vi.
  const jump = async (text) => {
    await page.getByRole("textbox", { name: "Nhảy tới ngày" }).fill(text);
    await page.getByRole("button", { name: "Đi tới" }).click();
    await settle();
  };
  await jump("06/08/2026");
  await waitDay("06/08/2026");
  await assertDayOnForm("06/08/2026");
  assert.equal((await cell("A9")).value.split("\n").length, 13);
  await jump("31/02/2026");
  assert.match(await lastMessage(), /không hợp lệ/);
  await jump("01/01/2027");
  assert.match(await lastMessage(), /không có trong vùng nguồn/);
  pass("Nhảy tới ngày", "06/08/2026 (13 dòng nội dung); báo lỗi 31/02/2026 và 01/01/2027");

  // 9. Mở lại task pane: cấu hình được khôi phục (R2.8); ô đích đã có dữ liệu → xin xác nhận (R2.5).
  const settings = await smoke(() => globalThis.__smoke.settings());
  const sheetsNow = await smoke(() => globalThis.__smoke.exportSheets());
  const browser = session.browser;
  await page.close();
  session = await openTaskpane(
    { sheets: sheetsNow, selection: { sheet: "Form", address: "A1" }, apiMinor, settings },
    { browser },
  );
  page = session.page;
  await page.getByRole("tab", { name: "Form mẫu" }).click();
  await page.waitForSelector("text=→ Form!A21");
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.waitForSelector("text=sẽ bị ghi đè");
  start = await logLength();
  await page.getByRole("button", { name: "Hủy" }).click();
  await settle();
  assert.equal((await writesSince(start)).length, 0);
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.getByRole("button", { name: "Ghi đè" }).click();
  await page.getByTestId("current-day").waitFor();
  await assertDayOnForm("06/08/2026");
  pass("Khôi phục cấu hình khi mở lại", "7 gán cột + ngày 06/08/2026; xác nhận ghi đè: Hủy không ghi, Ghi đè thì ghi");

  // 10. Ô đích chứa công thức → cảnh báo CÔNG THỨC; Hủy giữ nguyên công thức (R2.4).
  await smoke(() => globalThis.__smoke.select("Form", "A26:F26"));
  await page.getByRole("button", { name: "Gán ô đang chọn cho cột A" }).click();
  await page.waitForSelector("text=→ Form!A26");
  await page.getByRole("button", { name: "Xuất lại ngày đang chọn" }).click();
  await page.waitForSelector("text=CÔNG THỨC");
  await page.getByRole("button", { name: "Hủy" }).click();
  await settle();
  assert.equal((await cell("A26")).formula, fixture.sheets.find((s) => s.name === "Form").formulas.A26);
  await page.getByRole("button", { name: "Bỏ gán ô Form!A26" }).click();
  await settle();
  pass("Cảnh báo ô đích chứa công thức", "Hủy → công thức A26 giữ nguyên; bỏ gán A26");

  // 11. Sheet form bị đổi tên → báo lỗi, không ghi ô nào (R2.9).
  await smoke(() => globalThis.__smoke.renameSheet("Form", "Form_cu"));
  start = await logLength();
  await page.getByRole("button", { name: "Ngày sau" }).click();
  await settle();
  assert.match(await lastMessage(), /Không tìm thấy sheet: "Form"/);
  assert.equal((await writesSince(start)).length, 0);
  pass("Sheet form bị đổi tên", (await lastMessage()).trim());

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
