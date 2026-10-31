/*
 * Smoke test F1 trên bundle production với Excel giả lập.
 * Yêu cầu: đã chạy `npm run build`; có Playwright + Chromium trong môi trường (không phải dependency của dự án).
 *
 *   node scripts/smoke/run-f1.js <fixture.json> [--api 1] [--out <thư mục>] [--existing NK_TheoNgay]
 *                                [--delete-source]   (xóa sheet nguồn trước khi bấm Xuất)
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

function arg(name, fallback) {
  const i = process.argv.indexOf(name);
  return i === -1 ? fallback : process.argv[i + 1];
}

async function main() {
  const fixturePath = process.argv[2];
  const apiMinor = Number(arg("--api", "1"));
  const outDir = arg("--out", path.join(process.cwd(), "smoke-output"));
  const existing = arg("--existing", "");
  const dist = path.join(__dirname, "..", "..", "dist");
  const fixture = JSON.parse(fs.readFileSync(fixturePath, "utf8"));
  fs.mkdirSync(outDir, { recursive: true });

  const sheets = [{ name: fixture.sheetName, values: fixture.values }];
  if (existing) sheets.push({ name: existing, values: [["(sheet có sẵn)"]] });
  const smokeInput = { sheets, selection: { sheet: fixture.sheetName, address: fixture.address }, apiMinor };
  const officeStub =
    `window.__SMOKE_INPUT = ${JSON.stringify(smokeInput)};\n` +
    fs.readFileSync(path.join(__dirname, "fake-excel.js"), "utf8");

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 380, height: 1100 } });
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.route("**/*", (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "appsforoffice.microsoft.com") {
      return route.fulfill({ contentType: "application/javascript", body: officeStub });
    }
    if (url.hostname === "local.test") {
      const file = path.join(dist, decodeURIComponent(url.pathname));
      return route.fulfill({
        contentType: file.endsWith(".js") ? "application/javascript" : "text/html",
        body: fs.readFileSync(file),
      });
    }
    return route.abort();
  });

  await page.goto("https://local.test/taskpane.html");
  await page.getByRole("button", { name: "Lấy vùng đang chọn" }).click();
  await page.waitForSelector("text=Vùng input");
  await page.waitForTimeout(200);
  const panelText = await page.locator("#root").innerText();
  await page.screenshot({ path: path.join(outDir, "f1-preview.png"), fullPage: true });

  if (process.argv.includes("--delete-source")) {
    await page.evaluate((name) => globalThis.__smoke.deleteSheet(name), fixture.sheetName);
  }
  const exportButton = page.getByRole("button", { name: "Xuất nhật ký theo ngày" });
  let status = "(nút Xuất bị vô hiệu — không xuất)";

  if (await exportButton.isEnabled()) {
    await exportButton.click();
    await page.waitForSelector(".ms-MessageBar >> text=/Đã tạo sheet|Lỗi|Không/");
    status = await page.locator(".ms-MessageBar").last().innerText();
    const names = await page.evaluate(() => globalThis.__smoke.sheetNames());
    const created = names[names.length - 1];
    const result = { sheetName: created, rows: await page.evaluate((n) => globalThis.__smoke.dump(n), created) };
    fs.writeFileSync(path.join(outDir, "f1-result.json"), JSON.stringify(result, null, 1));
  }
  const log = await page.evaluate(() => globalThis.__smoke.log);
  await browser.close();

  console.log("=== PANEL ===\n" + panelText.replace(/\n{2,}/g, "\n"));
  console.log("=== STATUS ===\n" + status);
  console.log("=== WRITE LOG ===");
  for (const entry of log) console.log(" ", JSON.stringify(entry));
  console.log("=== PAGE ERRORS ===", pageErrors.length ? pageErrors : "none");
  if (pageErrors.length) process.exitCode = 1;
}

main().catch((error) => {
  console.error("SMOKE FAIL", error);
  process.exit(1);
});
