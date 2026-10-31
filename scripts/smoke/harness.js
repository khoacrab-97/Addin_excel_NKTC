/*
 * Khởi động task pane (bundle production trong dist/) với Excel giả lập trong Chromium.
 * Dùng chung cho các smoke test. Yêu cầu Playwright + Chromium có sẵn trong môi trường.
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const DIST = path.join(__dirname, "..", "..", "dist");

/**
 * @param {{ sheets: object[], selection: { sheet: string, address: string }, apiMinor?: number, settings?: object }} input
 * @param {{ browser?: import("playwright").Browser, viewport?: { width: number, height: number },
 *   overrides?: Record<string, { status?: number, body: string } | "abort"> }} [options]
 *   overrides: thay phản hồi cho đường dẫn trong dist (ví dụ "/version.json") — giả lập phiên bản mới / mất mạng.
 */
async function openTaskpane(input, options = {}) {
  const browser = options.browser ?? (await chromium.launch());
  const page = await browser.newPage({ viewport: options.viewport ?? { width: 380, height: 1100 } });
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  const officeStub =
    `window.__SMOKE_INPUT = ${JSON.stringify({ apiMinor: 1, ...input })};\n` +
    fs.readFileSync(path.join(__dirname, "fake-excel.js"), "utf8");
  await page.route("**/*", (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "appsforoffice.microsoft.com") {
      return route.fulfill({ contentType: "application/javascript", body: officeStub });
    }
    if (url.hostname === "local.test") {
      const override = options.overrides?.[url.pathname];
      if (override === "abort") return route.abort();
      if (override)
        return route.fulfill({ status: override.status ?? 200, contentType: "application/json", body: override.body });
      const file = path.join(DIST, decodeURIComponent(url.pathname));
      return route.fulfill({
        contentType: file.endsWith(".js") ? "application/javascript" : "text/html",
        body: fs.readFileSync(file),
      });
    }
    return route.abort();
  });
  await page.goto("https://local.test/taskpane.html");
  return { browser, page, pageErrors };
}

module.exports = { openTaskpane };
