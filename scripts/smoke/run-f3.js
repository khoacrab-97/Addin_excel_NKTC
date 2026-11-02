/*
 * Smoke test F3 (nhóm Tiện ích trên Ribbon): hướng dẫn, file tải về, phiên bản + kiểm tra cập nhật;
 * trang chào mừng khi kích hoạt add-in.
 *   node scripts/smoke/run-f3.js
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { chromium } = require("playwright");
const { openWorkbook } = require("./harness");

const DIST = path.join(__dirname, "..", "..", "dist");
const pkg = require("../../package.json");
const INPUT = { sheets: [{ name: "Sheet1", values: [[""]] }], selection: { sheet: "Sheet1", address: "A1" } };

const steps = [];
const pass = (name, detail = "") => steps.push(`  ✔ ${name}${detail ? ` — ${detail}` : ""}`);

async function main() {
  const browser = await chromium.launch();
  const errors = [];
  const version = (v) => ({ body: JSON.stringify({ version: v, builtAt: "2026-11-02T00:00:00Z" }) });

  /** Mở hộp thoại Phiên bản với version.json cho trước; trả về nội dung khi có kết quả kiểm tra. */
  async function about(overrides, waitText) {
    const wb = await openWorkbook(INPUT, { browser, overrides });
    const done = await wb.press("nktcAbout");
    const d = await wb.dialog();
    await d.page.waitForSelector(`text=${waitText}`);
    const text = await d.text();
    const screenshot = d.page.screenshot.bind(d.page);
    await d.click("Đóng");
    await done();
    errors.push(...wb.pageErrors);
    await wb.context.close();
    return { text, screenshot };
  }

  // 1. Hướng dẫn + file tải về có trong bản build (P3.1).
  const wb = await openWorkbook(INPUT, { browser });
  const done = await wb.press("nktcHelp");
  const d = await wb.dialog();
  const text = await d.text();
  assert.ok(text.includes("tab Nhật ký thi công") && text.includes("Xuất theo ngày") && text.includes("Gán cột"));
  const hrefs = await d.page.locator("a").evaluateAll((links) => links.map((a) => a.getAttribute("href")));
  for (const relative of ["templates/Mau_NKTC.xlsx", "templates/Mau_TienDo.xlsx"]) {
    assert.ok(hrefs.includes(relative), `Thiếu link ${relative}`);
    assert.ok(fs.statSync(path.join(DIST, relative)).size > 0, `Thiếu file dist/${relative}`);
  }
  assert.ok(hrefs.some((h) => h?.includes("HUONG_DAN_CAI_DAT.md")));
  await d.page.screenshot({ path: path.join("smoke-output", "f3-help.png"), fullPage: true });
  await d.click("Đóng");
  await done();
  errors.push(...wb.pageErrors);
  await wb.context.close();
  pass("Hướng dẫn (Ribbon)", "nội dung theo nút Ribbon; Mau_NKTC.xlsx, Mau_TienDo.xlsx có trong dist");

  // 2. Phiên bản: bản đang chạy = bản phát hành → "mới nhất" (P3.2).
  const latest = await about(undefined, "phiên bản mới nhất");
  assert.ok(latest.text.includes(`Phiên bản add-in: ${pkg.version}`) && latest.text.includes("ExcelApi tối đa: 1.1"));
  pass("Phiên bản — bản mới nhất", `version.json thật của bản build (${pkg.version}), ExcelApi 1.1`);

  // 3. Máy chủ có bản mới → hướng dẫn cập nhật (P3.3).
  const newer = await about({ "/version.json": version("9.9.9") }, "Đã có phiên bản");
  assert.match(newer.text, /9\.9\.9/);
  assert.ok(newer.text.includes("đóng hẳn Excel rồi mở lại"));
  pass("Có bản mới", "9.9.9 → các bước cập nhật");

  // 4. Bản đang chạy mới hơn bản phát hành (môi trường phát triển).
  await about({ "/version.json": version("0.0.1") }, "bản phát triển");
  pass("Bản phát triển mới hơn bản phát hành", "không báo cập nhật nhầm");

  // 5. Lỗi: mất mạng, HTTP 404, nội dung hỏng → báo không kiểm tra được, không lỗi JS.
  for (const [label, override] of [
    ["mất mạng", "abort"],
    ["HTTP 404", { status: 404, body: "Not Found" }],
    ["nội dung hỏng", { body: "<html>oops</html>" }],
  ]) {
    await about({ "/version.json": override }, "Không kiểm tra được phiên bản mới");
    pass(`Lỗi kiểm tra cập nhật: ${label}`, "báo lỗi thân thiện");
  }

  // 6. Trang chào mừng (khung hiện một lần khi bộ cài kích hoạt add-in).
  const welcome = await openWorkbook(INPUT, { browser });
  const page = await welcome.context.newPage();
  page.on("pageerror", (e) => errors.push(`welcome: ${e.message}`));
  await page.goto("https://local.test/taskpane.html");
  await page.waitForFunction((v) => globalThis.document.getElementById("version")?.textContent === v, pkg.version);
  assert.match(await page.locator("body").innerText(), /tab Nhật ký thi công/);
  errors.push(...welcome.pageErrors);
  pass("Trang chào mừng", `hướng dẫn dùng tab Ribbon, phiên bản ${pkg.version}`);

  // 7. Hộp thoại mở trực tiếp không có dữ liệu → thông báo, không lỗi JS.
  const bad = await welcome.context.newPage();
  bad.on("pageerror", (e) => errors.push(`dialog: ${e.message}`));
  await bad.goto("https://local.test/dialog.html?view=assign&id=none");
  await bad.waitForSelector("text=Không đọc được dữ liệu của hộp thoại");
  pass("Hộp thoại thiếu dữ liệu", "báo lỗi, không chạy tiếp");

  await browser.close();
  console.log(`Smoke F3 — ${steps.length} bước:\n${steps.join("\n")}`);
  console.log("PAGE ERRORS:", errors.length ? errors : "none");
  if (errors.length) process.exitCode = 1;
}

fs.mkdirSync("smoke-output", { recursive: true });
main().catch((error) => {
  console.log(steps.join("\n"));
  console.error("SMOKE F3 FAIL:", error.message);
  process.exit(1);
});
