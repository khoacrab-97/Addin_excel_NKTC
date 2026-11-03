/*
 * Chạy add-in (bundle production trong dist/) với Excel giả lập trong Chromium — kiểu Ribbon:
 * trang commands.html (function file) nhận lệnh của nút; hộp thoại dialog.html mở ở trang riêng
 * trong cùng BrowserContext (dùng chung localStorage như Excel). Yêu cầu Playwright + Chromium có sẵn.
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const DIST = path.join(__dirname, "..", "..", "dist");

function watchErrors(page, pageErrors, label) {
  page.on("pageerror", (error) => pageErrors.push(`${label}: ${error.message}`));
  // Vi phạm Content-Security-Policy được trình duyệt báo qua console → coi là lỗi.
  page.on("console", (message) => {
    if (/Content Security Policy|Refused to (load|execute|apply)/i.test(message.text()))
      pageErrors.push(`${label}: ${message.text()}`);
  });
}

/**
 * @param {{ sheets: object[], selection: { sheet: string, address: string }, apiMinor?: number, settings?: object }} input
 * @param {{ browser?: import("playwright").Browser,
 *   overrides?: Record<string, { status?: number, body: string } | "abort"> }} [options]
 *   overrides: thay phản hồi cho đường dẫn trong dist (ví dụ "/version.json") — giả lập phiên bản mới / mất mạng.
 */
async function openWorkbook(input, options = {}) {
  const browser = options.browser ?? (await chromium.launch());
  const context = await browser.newContext({ viewport: { width: 640, height: 720 } });
  const pageErrors = [];
  const officeStub =
    `window.__SMOKE_INPUT = ${JSON.stringify({ apiMinor: 1, ...input })};\n` +
    fs.readFileSync(path.join(__dirname, "fake-excel.js"), "utf8");
  await context.route("**/*", (route) => {
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
        contentType: file.endsWith(".js")
          ? "application/javascript"
          : file.endsWith(".png")
            ? "image/png"
            : "text/html",
        body: fs.readFileSync(file),
      });
    }
    return route.abort();
  });

  const commands = await context.newPage();
  watchErrors(commands, pageErrors, "commands");
  /** Trang hộp thoại đang mở — nhận tin messageChild (DialogApi 1.2) từ function file. */
  let currentDialogPage;
  await commands.exposeFunction("__toChild", async (message) => {
    if (currentDialogPage && !currentDialogPage.isClosed()) {
      await currentDialogPage.evaluate((m) => globalThis.__parentHandler?.({ message: m }), message);
    }
  });
  await commands.goto("https://local.test/commands.html");
  await commands.waitForFunction(() => typeof globalThis.nktcExportDaily === "function");

  const smoke = (fn, arg) => commands.evaluate(fn, arg);
  let runCount = 0;

  /** Bấm nút trên Ribbon (gọi hàm ExecuteFunction). Trả về hàm chờ lệnh hoàn tất (event.completed). */
  async function press(name) {
    const id = ++runCount;
    await commands.evaluate(
      ([fn, runId]) => {
        globalThis.__done = globalThis.__done || [];
        globalThis[fn]({
          completed: () => {
            globalThis.__done.push(runId);
          },
        });
      },
      [name, id],
    );
    return () => commands.waitForFunction((runId) => (globalThis.__done || []).includes(runId), id);
  }

  /** Chờ hộp thoại mở, hiển thị dialog.html ở trang riêng. */
  async function dialog() {
    await commands.waitForFunction(() => globalThis.__smoke.openDialog() !== null);
    const info = await smoke(() => globalThis.__smoke.openDialog());
    const page = await context.newPage();
    currentDialogPage = page;
    watchErrors(page, pageErrors, "dialog");
    await page.exposeFunction("__toParent", (message) =>
      commands.evaluate((m) => globalThis.__smoke.deliver(m), message),
    );
    await page.goto(info.url);
    await page.waitForFunction(() => (globalThis.document.getElementById("root")?.innerText ?? "").trim() !== "");
    /** Bấm nút trong hộp thoại; nút gửi kết quả → hộp thoại đóng. */
    const click = async (name) => {
      await page.getByRole("button", { name, exact: true }).click();
      await commands.waitForFunction((index) => globalThis.__smoke.dialogClosed(index), info.index);
      await page.close();
    };
    /** Thao tác bất kỳ trong hộp thoại làm hộp thoại gửi kết quả (ví dụ nhấn Enter, nút ⬆) rồi chờ đóng. */
    const act = async (action) => {
      await action(page);
      await commands.waitForFunction((index) => globalThis.__smoke.dialogClosed(index), info.index);
      await page.close();
    };
    const text = () => page.locator("#root").innerText();
    /** Chọn giá trị trong ô chọn cột (Fluent Dropdown). */
    const choose = async (label, optionPrefix) => {
      await page.getByRole("combobox", { name: label }).click();
      await page
        .getByRole("option")
        .filter({ hasText: new RegExp(`^${optionPrefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`) })
        .first()
        .click();
    };
    /** Bấm nút mà hộp thoại vẫn mở (bảng chuyển ngày, DialogApi 1.2). */
    const press = (name) => page.getByRole("button", { name, exact: true }).click();
    /** Đóng trang hộp thoại khi add-in đã đóng hộp thoại (ví dụ bảng chuyển ngày trên bản cũ / khi hỏi xác nhận). */
    const closedByAddin = async () => {
      await commands.waitForFunction((index) => globalThis.__smoke.dialogClosed(index), info.index);
      if (!page.isClosed()) await page.close();
    };
    return { page, info, click, act, press, closedByAddin, text, choose };
  }

  /**
   * Nút ⬆ của ô chọn vùng (giống Excel): hộp thoại chính đóng → thanh chọn (tiêu đề + ô nhập + nút ⬇) → chọn trên
   * sheet → địa chỉ hiện ngay trong ô nhập → bấm ⬇ / Enter → hộp thoại chính mở lại.
   * `select` = [sheet, address] cần chọn (bỏ qua → không chọn lại). Tùy chọn: `shown` = địa chỉ phải hiện trong
   * thanh sau khi chọn; `type` = gõ địa chỉ vào thanh; `confirm` = "arrow" | "enter" | "close" (X của cửa sổ);
   * `inspect(bar)` = kiểm tra thêm khi thanh đang mở.
   */
  async function pickVia(d, buttonName, select, options = {}) {
    const confirm = options.confirm ?? "arrow";
    await d.act((page) => page.getByRole("button", { name: buttonName }).click());
    const bar = await dialog();
    const box = bar.page.getByRole("textbox", { name: "Địa chỉ đang chọn" });
    if (options.initial !== undefined) {
      const initial = await box.inputValue();
      if (initial !== options.initial) throw new Error(`Thanh chọn mở với "${initial}", cần "${options.initial}"`);
    }
    if (select) await smoke(([s, a]) => globalThis.__smoke.select(s, a), select);
    if (options.shown !== undefined) {
      await bar.page.waitForFunction(
        (want) => globalThis.document.querySelector('input[aria-label="Địa chỉ đang chọn"]')?.value === want,
        options.shown,
      );
    }
    if (options.type !== undefined) await box.fill(options.type);
    if (options.inspect) await options.inspect(bar);
    if (confirm === "close") {
      await smoke(() => globalThis.__smoke.dismiss());
      await bar.page.close();
    } else if (confirm === "enter") {
      await bar.act(() => box.press("Enter"));
    } else {
      await bar.click("Xác nhận vùng chọn");
    }
    return dialog();
  }

  return { browser, context, commands, pageErrors, smoke, press, dialog, pickVia };
}

module.exports = { openWorkbook };
