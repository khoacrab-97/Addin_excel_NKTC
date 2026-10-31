/*
 * Smoke test F3 (tab Tiện ích): hướng dẫn, file tải về, kiểm tra cập nhật, thông báo bản mới.
 *   node scripts/smoke/run-f3.js
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert/strict");
const { chromium } = require("playwright");
const { openTaskpane } = require("./harness");

const DIST = path.join(__dirname, "..", "..", "dist");
const pkg = require("../../package.json");
const INPUT = { sheets: [{ name: "Sheet1", values: [[""]] }], selection: { sheet: "Sheet1", address: "A1" } };

const steps = [];
const pass = (name, detail = "") => steps.push(`  ✔ ${name}${detail ? ` — ${detail}` : ""}`);

async function main() {
  const browser = await chromium.launch();
  const errors = [];
  const open = async (overrides) => {
    const session = await openTaskpane(INPUT, { browser, overrides });
    session.page.on("pageerror", (e) => errors.push(e.message));
    return session.page;
  };
  const version = (v) => ({ body: JSON.stringify({ version: v, builtAt: "2026-10-31T00:00:00Z" }) });

  // 1. Hướng dẫn + file tải về có trong bản build (P3.1).
  let page = await open();
  await page.getByRole("tab", { name: "Tiện ích" }).click();
  await page.waitForSelector("text=Hướng dẫn sử dụng");
  const hrefs = await page.locator("a").evaluateAll((links) => links.map((a) => a.getAttribute("href")));
  for (const relative of ["templates/Mau_NKTC.xlsx", "templates/Mau_TienDo.xlsx"]) {
    assert.ok(hrefs.includes(relative), `Thiếu link ${relative}`);
    assert.ok(fs.statSync(path.join(DIST, relative)).size > 0, `Thiếu file dist/${relative}`);
  }
  assert.ok(hrefs.some((h) => h?.includes("HUONG_DAN_CAI_DAT.md")));
  pass("Hướng dẫn F1/F2 và link tải", "Mau_NKTC.xlsx, Mau_TienDo.xlsx có trong dist; link hướng dẫn cài đặt");

  // 2. Bản đang chạy = bản phát hành → "mới nhất"; không có thông báo đầu trang (P3.2).
  await page.getByRole("button", { name: "Kiểm tra cập nhật" }).click();
  await page.waitForSelector("text=phiên bản mới nhất");
  assert.equal(await page.locator("text=Xem hướng dẫn cập nhật").count(), 0);
  pass("Kiểm tra cập nhật — bản mới nhất", `version.json thật của bản build (${pkg.version})`);
  await page.screenshot({ path: path.join("smoke-output", "f3-utilities.png"), fullPage: true });
  await page.close();

  // 3. Máy chủ có bản mới → thông báo đầu trang ở tab khác, nút dẫn tới hướng dẫn cập nhật (P3.3).
  page = await open({ "/version.json": version("9.9.9") });
  await page.waitForSelector("text=Đã có phiên bản mới 9.9.9");
  await page.screenshot({ path: path.join("smoke-output", "f3-banner.png") });
  await page.getByRole("button", { name: "Xem hướng dẫn cập nhật" }).click();
  await page.waitForSelector("text=Đóng khung add-in rồi mở lại");
  assert.equal(await page.getByRole("tab", { name: "Tiện ích" }).getAttribute("aria-selected"), "true");
  pass("Thông báo bản mới", "banner 9.9.9 → nút chuyển sang tab Tiện ích, hiện các bước cập nhật");
  await page.close();

  // 4. Bản đang chạy mới hơn bản phát hành (môi trường phát triển).
  page = await open({ "/version.json": version("0.0.1") });
  await page.getByRole("tab", { name: "Tiện ích" }).click();
  await page.waitForSelector("text=bản phát triển");
  pass("Bản phát triển mới hơn bản phát hành", "không báo cập nhật nhầm");
  await page.close();

  // 5. Lỗi: mất mạng, HTTP 404, nội dung hỏng → báo không kiểm tra được, không lỗi JS.
  for (const [label, override] of [
    ["mất mạng", "abort"],
    ["HTTP 404", { status: 404, body: "Not Found" }],
    ["nội dung hỏng", { body: "<html>oops</html>" }],
  ]) {
    page = await open({ "/version.json": override });
    await page.getByRole("tab", { name: "Tiện ích" }).click();
    await page.waitForSelector("text=Không kiểm tra được phiên bản mới");
    await page.close();
    pass(`Lỗi kiểm tra cập nhật: ${label}`, "báo lỗi thân thiện");
  }

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
