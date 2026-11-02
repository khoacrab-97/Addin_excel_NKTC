/*
 * Smoke test F1 (nút "Xuất theo ngày" trên Ribbon) trên bundle production với Excel giả lập.
 * Yêu cầu: đã chạy `npm run build`; có Playwright + Chromium trong môi trường.
 *
 *   node scripts/smoke/run-f1.js <tên file trong test/fixtures> [--api 1] [--existing NK_TheoNgay]
 * Log chỉ in tên bước và số liệu, không in nội dung dữ liệu test.
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { readFixture } = require("./fixtures");
const { openWorkbook } = require("./harness");

function arg(name, fallback) {
  const i = process.argv.indexOf(name);
  return i === -1 ? fallback : process.argv[i + 1];
}

const steps = [];
const pass = (name, detail = "") => steps.push(`  ✔ ${name}${detail ? ` — ${detail}` : ""}`);

async function main() {
  const fixture = readFixture(process.argv[2] ?? "tdtc-real.json");
  const apiMinor = Number(arg("--api", "1"));
  const existing = arg("--existing", "");
  const outDir = path.join(process.cwd(), "smoke-output");
  fs.mkdirSync(outDir, { recursive: true });

  const sheets = [{ name: fixture.sheetName, values: fixture.values }];
  if (existing) sheets.push({ name: existing, values: [["(sheet có sẵn)"]] });
  // Chỉ đặt con trỏ vào 1 ô → add-in lấy cả vùng có dữ liệu của sheet.
  const wb = await openWorkbook({ sheets, selection: { sheet: fixture.sheetName, address: "B5" }, apiMinor });
  const sheetNames = () => wb.smoke(() => globalThis.__smoke.sheetNames());

  // 1. Hộp thoại xác nhận cột: nhận diện sẵn 3 cột, vùng = used range.
  let done = await wb.press("nktcExportDaily");
  let d = await wb.dialog();
  let text = await d.text();
  assert.ok(text.includes(`Vùng dữ liệu: ${fixture.sheetName}!${fixture.address}`), text);
  for (const label of ["Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"]) {
    const value = await d.page.getByRole("combobox", { name: label }).innerText();
    assert.ok(!value.includes("Chọn cột"), `${label} chưa được nhận diện`);
  }
  assert.ok(await d.page.getByRole("button", { name: "Xuất nhật ký" }).isEnabled());
  await d.page.screenshot({ path: path.join(outDir, "f1-columns.png") });
  pass("Hộp thoại chọn cột", "đặt con trỏ 1 ô → lấy cả vùng dữ liệu; 3 cột nhận diện sẵn");

  // 2. Xuất → sheet mới không trùng tên, thông báo kết quả + xem trước 5 ngày.
  await d.click("Xuất nhật ký");
  d = await wb.dialog();
  text = await d.text();
  const created = (await sheetNames()).at(-1);
  assert.equal(created, existing ? `${existing}_2` : "NK_TheoNgay");
  assert.ok(text.includes(`Đã tạo sheet "${created}" gồm `), text);
  assert.ok(text.includes("công việc hợp lệ"));
  assert.equal(await d.page.locator("tbody tr").count(), 5);
  await d.page.screenshot({ path: path.join(outDir, "f1-result.png") });
  const rows = await wb.smoke((n) => globalThis.__smoke.dump(n), created);
  assert.deepEqual([rows[0][0].value, rows[0][1].value], ["Ngày", "Nội dung"]);
  assert.equal(rows[1][0].numberFormat, "dd/mm/yyyy");
  fs.writeFileSync(path.join(outDir, "f1-result.json"), JSON.stringify({ sheetName: created, rows }, null, 1));
  await d.click("Đóng");
  await done();
  pass("Xuất nhật ký", `sheet mới không trùng tên, ${rows.length - 1} ngày, xem trước 5 ngày`);

  // 3. Hủy → không tạo sheet.
  const before = (await sheetNames()).length;
  done = await wb.press("nktcExportDaily");
  d = await wb.dialog();
  await d.click("Hủy");
  await done();
  assert.equal((await sheetNames()).length, before);
  pass("Hủy trong hộp thoại", "không tạo sheet");

  // 4. Đóng hộp thoại bằng nút X (DialogEventReceived 12006) → lệnh kết thúc, không tạo sheet.
  done = await wb.press("nktcExportDaily");
  await wb.commands.waitForFunction(() => globalThis.__smoke.openDialog() !== null);
  await wb.smoke(() => globalThis.__smoke.dismiss());
  await done();
  assert.equal((await sheetNames()).length, before);
  pass("Đóng hộp thoại bằng nút X", "lệnh kết thúc, không tạo sheet");

  // 5. Vùng chỉ có 1 dòng → cảnh báo, không mở hộp thoại chọn cột.
  await wb.smoke(() => globalThis.__smoke.select("TDTC", "A1:C1"));
  done = await wb.press("nktcExportDaily");
  d = await wb.dialog();
  assert.match(await d.text(), /cần có dòng tiêu đề và ít nhất một dòng công việc/);
  await d.click("Đóng");
  await done();
  pass("Vùng thiếu dữ liệu", "cảnh báo hướng dẫn chọn lại");

  const log = await wb.smoke(() => globalThis.__smoke.log);
  const outside = log.filter((e) => e.op === "values" && !e.range.startsWith(created));
  assert.deepEqual(outside, [], "Ghi ngoài sheet kết quả");
  await wb.browser.close();
  console.log(`Smoke F1 (ExcelApi 1.${apiMinor}) — ${steps.length} bước:\n${steps.join("\n")}`);
  console.log("PAGE ERRORS:", wb.pageErrors.length ? wb.pageErrors : "none");
  if (wb.pageErrors.length) process.exitCode = 1;
}

main().catch((error) => {
  console.log(steps.join("\n"));
  console.error("SMOKE F1 FAIL:", error.message);
  process.exit(1);
});
