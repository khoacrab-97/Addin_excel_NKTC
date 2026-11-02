/*
 * Smoke test F4 (nút "Tạo bảng" trên Ribbon) trên bundle production với Excel giả lập:
 * tạo bảng trên sheet mới / sheet hiện tại, không ghi đè vùng có dữ liệu, rồi "Xuất theo ngày" từ bảng vừa điền.
 *   node scripts/smoke/run-f4.js
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { chromium } = require("playwright");
const { openWorkbook } = require("./harness");

const HEADERS = ["STT", "Nội dung công việc", "Bắt đầu", "Kết thúc"];
const steps = [];
const pass = (name, detail = "") => steps.push(`  ✔ ${name}${detail ? ` — ${detail}` : ""}`);
/** Ngày Excel (serial, hệ 1900) từ dd/mm/yyyy. */
const serial = (text) => {
  const [d, m, y] = text.split("/").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
};

async function main() {
  const browser = await chromium.launch();
  const wb = await openWorkbook(
    {
      sheets: [
        { name: "Sheet1", values: [["Ghi chú có sẵn", "", "", "", "", ""]] },
        { name: "TienDo", values: [["(sheet có sẵn)"]] },
      ],
      selection: { sheet: "Sheet1", address: "B3" },
    },
    { browser },
  );
  const dump = (name) => wb.smoke((n) => globalThis.__smoke.dump(n), name);
  const tables = (name) => wb.smoke((n) => globalThis.__smoke.tables(n), name);
  const set = (sheet, address, value) =>
    wb.smoke(([s, a, v]) => globalThis.__smoke.setValue(s, a, v), [sheet, address, value]);

  /** Bấm "Tạo bảng", thao tác trong hộp thoại; trả về hàm chờ lệnh hoàn tất. */
  async function createTable(fill) {
    const finished = await wb.press("nktcCreateTable");
    await fill(await wb.dialog());
    return finished;
  }

  // 1. Mặc định: sheet mới, 10 dòng → sheet TienDo_2 (không trùng), bảng A1:D11.
  let done = await createTable(async (dialog) => {
    const text = await dialog.text();
    assert.ok(text.includes("STT | Nội dung công việc | Bắt đầu | Kết thúc"), text);
    assert.equal(await dialog.page.getByRole("textbox", { name: /Số dòng/ }).inputValue(), "10");
    assert.match(text, /Bảng sẽ chiếm vùng A1:D11/);
    await dialog.page.screenshot({ path: path.join("smoke-output", "f4-dialog.png") });
    await dialog.click("Tạo bảng");
  });
  let d = await wb.dialog();
  assert.match(await d.text(), /Đã tạo bảng TienDo_2!A1:D11 gồm 10 dòng/);
  await d.click("Đóng");
  await done();
  let rows = await dump("TienDo_2");
  assert.deepEqual(
    rows[0].map((cell) => cell.value),
    HEADERS,
  );
  assert.deepEqual(
    rows.slice(1).map((row) => row[0].value),
    [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
  );
  assert.equal(rows[1][1].numberFormat, "@");
  assert.equal(rows[1][2].numberFormat, "dd/mm/yyyy");
  assert.equal(rows[10][3].numberFormat, "dd/mm/yyyy");
  assert.deepEqual(await tables("TienDo_2"), [{ name: "Table1", address: "A1:D11", style: "TableStyleMedium2" }]);
  assert.equal(await wb.smoke(() => globalThis.__smoke.activeSheet()), "TienDo_2");
  assert.deepEqual(await wb.smoke(() => globalThis.__smoke.selection()), { sheet: "TienDo_2", address: "B2:B2" });
  pass(
    "Sheet mới, mặc định 10 dòng",
    "TienDo_2!A1:D11 (TienDo có sẵn), STT 1..10, định dạng ngày, Excel Table, con trỏ ở B2",
  );

  // 2. Sheet hiện tại, ô bắt đầu B3, 5 dòng → B3:E8; ô ghi chú A1 giữ nguyên.
  await wb.smoke(() => globalThis.__smoke.select("Sheet1", "B3"));
  done = await createTable(async (dialog) => {
    await dialog.page.getByText("Sheet hiện tại (Sheet1)").click();
    const cellBox = dialog.page.getByRole("textbox", { name: /Ô bắt đầu/ });
    assert.equal(await cellBox.inputValue(), "B3");
    await dialog.page.getByRole("textbox", { name: /Số dòng/ }).fill("5");
    assert.match(await dialog.text(), /Bảng sẽ chiếm vùng B3:E8 — vùng này phải đang trống/);
    await dialog.click("Tạo bảng");
  });
  d = await wb.dialog();
  assert.match(await d.text(), /Đã tạo bảng Sheet1!B3:E8 gồm 5 dòng/);
  await d.click("Đóng");
  await done();
  rows = await dump("Sheet1");
  assert.equal(rows[0][0].value, "Ghi chú có sẵn");
  assert.deepEqual(
    rows[2].slice(1, 5).map((cell) => cell.value),
    HEADERS,
  );
  assert.equal(rows[7][1].value, 5);
  pass("Sheet hiện tại, ô B3, 5 dòng", "Sheet1!B3:E8; dữ liệu có sẵn ở A1 giữ nguyên");

  // 3. Vùng đích có dữ liệu → báo lỗi, không ghi gì (R4.4).
  const before = (await wb.smoke(() => globalThis.__smoke.log.length)) ?? 0;
  await wb.smoke(() => globalThis.__smoke.select("Sheet1", "A1"));
  done = await createTable(async (dialog) => {
    await dialog.page.getByText("Sheet hiện tại (Sheet1)").click();
    await dialog.click("Tạo bảng");
  });
  d = await wb.dialog();
  const error = await d.text();
  assert.match(error, /Vùng A1:D11 đang có dữ liệu \(A1, B3, C3, D3/);
  assert.match(error, /không ghi đè/);
  await d.click("Đóng");
  await done();
  const writes = await wb.smoke(
    (start) => globalThis.__smoke.log.slice(start).filter((e) => e.op === "values" || e.op === "tables.add"),
    before,
  );
  assert.deepEqual(writes, []);
  pass("Vùng đích đã có dữ liệu", "báo các ô có dữ liệu, không ghi ô nào");

  // 4. Số dòng / ô sai → nút Tạo bảng bị khóa; Hủy → không tạo gì.
  done = await createTable(async (dialog) => {
    await dialog.page.getByText("Sheet hiện tại (Sheet1)").click();
    const create = dialog.page.getByRole("button", { name: "Tạo bảng", exact: true });
    await dialog.page.getByRole("textbox", { name: /Số dòng/ }).fill("0");
    assert.ok(await create.isDisabled());
    await dialog.page.getByRole("textbox", { name: /Số dòng/ }).fill("1001");
    assert.ok(await create.isDisabled());
    await dialog.page.getByRole("textbox", { name: /Số dòng/ }).fill("10");
    await dialog.page.getByRole("textbox", { name: /Ô bắt đầu/ }).fill("B");
    assert.ok(await create.isDisabled());
    await dialog.click("Hủy");
  });
  await done();
  assert.equal((await wb.smoke(() => globalThis.__smoke.sheetNames())).length, 3);
  pass("Kiểm tra nhập liệu", "0 / 1001 dòng, ô 'B' → khóa nút; Hủy không tạo");

  // 5. Điền bảng vừa tạo → "Xuất theo ngày" nhận đủ 3 cột và xuất đúng.
  await set("TienDo_2", "B2", "Đào đất hố móng");
  await set("TienDo_2", "C2", serial("01/11/2026"));
  await set("TienDo_2", "D2", serial("03/11/2026"));
  await set("TienDo_2", "B3", "Đổ bê tông lót");
  await set("TienDo_2", "C3", "03/11/2026");
  await set("TienDo_2", "D3", "04/11/2026");
  await wb.smoke(() => globalThis.__smoke.select("TienDo_2", "B2"));
  done = await wb.press("nktcExportDaily");
  d = await wb.dialog();
  for (const [label, column] of [
    ["Nội dung công việc", "B – Nội dung công việc"],
    ["Ngày bắt đầu", "C – Bắt đầu"],
    ["Ngày kết thúc", "D – Kết thúc"],
  ]) {
    assert.ok((await d.page.getByRole("combobox", { name: label }).innerText()).startsWith(column), label);
  }
  await d.click("Xuất nhật ký");
  d = await wb.dialog();
  const result = await d.text();
  assert.match(result, /gồm 4 ngày, từ 01\/11\/2026 đến 04\/11\/2026/);
  assert.match(result, /2 công việc hợp lệ/);
  assert.doesNotMatch(result, /dòng lỗi/);
  await d.click("Đóng");
  await done();
  pass("Điền bảng → Xuất theo ngày", "3 cột tự nhận diện; 2 công việc → 4 ngày; dòng trống trong bảng không báo lỗi");

  const errors = wb.pageErrors;
  await browser.close();
  console.log(`Smoke F4 — ${steps.length} bước:\n${steps.join("\n")}`);
  console.log("PAGE ERRORS:", errors.length ? errors : "none");
  if (errors.length) process.exitCode = 1;
}

fs.mkdirSync("smoke-output", { recursive: true });
main().catch((error) => {
  console.log(steps.join("\n"));
  console.error("SMOKE F4 FAIL:", error.message);
  process.exit(1);
});
