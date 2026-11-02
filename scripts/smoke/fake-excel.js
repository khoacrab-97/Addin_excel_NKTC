/*
 * Excel/Office giả lập (chạy trong trình duyệt) cho smoke test. KHÔNG dùng trong add-in.
 * Mô phỏng một phần Office.js ExcelApi 1.1 đủ cho các luồng của add-in:
 * workbook trong bộ nhớ, getRange/getUsedRange/getIntersection/getSelectedRange,
 * ghi values/numberFormat/format, lỗi ItemNotFound được trả về lúc context.sync().
 * Đầu vào: window.__SMOKE_INPUT = {
 *   sheets: [{ name, values, formulas?: { "A26": "=..." } }], selection: { sheet, address }, apiMinor, settings? }.
 * Ghi chuỗi bắt đầu bằng "'" → lưu phần sau dấu "'" (giống Excel: dấu "'" là ký tự đánh dấu văn bản).
 * Dialog API 1.1 giả lập: displayDialogAsync ghi nhận hộp thoại; test mở URL hộp thoại ở trang khác và chuyển
 * messageParent về qua window.__toParent → __smoke.deliver. Không có Office.actions (giống Excel 2016).
 */
(function () {
  "use strict";
  var MAX_ROW = 1048575;
  var input = window.__SMOKE_INPUT;
  var log = [];
  var sheets = [];
  var selection = { sheet: input.selection.sheet, address: input.selection.address };
  var settingsStore = JSON.parse(JSON.stringify(input.settings || {}));
  var dialogs = [];
  function openDialog() {
    for (var i = dialogs.length - 1; i >= 0; i--) if (!dialogs[i].closed) return dialogs[i];
    return null;
  }

  function colIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function colLetters(index) {
    var s = "";
    var n = index + 1;
    while (n > 0) {
      s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function parseAddress(address) {
    var local = address.slice(address.lastIndexOf("!") + 1).toUpperCase();
    var parts = local.split(":");
    function parsePart(part) {
      var m = /^\$?([A-Z]*)\$?(\d*)$/.exec(part);
      if (!m) throw { code: "InvalidArgument", message: "Địa chỉ không hợp lệ: " + address };
      return { c: m[1] ? colIndex(m[1]) : null, r: m[2] ? Number(m[2]) - 1 : null };
    }
    var a = parsePart(parts[0]);
    var b = parsePart(parts[1] || parts[0]);
    return {
      r1: a.r === null ? 0 : a.r,
      c1: a.c === null ? 0 : a.c,
      r2: b.r === null ? MAX_ROW : b.r,
      c2: b.c === null ? 16383 : b.c,
    };
  }
  function quoteSheet(name) {
    return /^[A-Za-z0-9_]+$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
  }

  function Sheet(name) {
    this.name = name;
    this.cells = {};
  }
  Sheet.prototype.cell = function (r, c) {
    var key = r + "," + c;
    if (!this.cells[key]) this.cells[key] = { value: "", formula: null, numberFormat: "General", format: {} };
    return this.cells[key];
  };
  Sheet.prototype.peek = function (r, c) {
    return this.cells[r + "," + c];
  };
  Sheet.prototype.usedBox = function () {
    var box = null;
    for (var key in this.cells) {
      if (this.cells[key].value === "" && !this.cells[key].formula) continue;
      var rc = key.split(",").map(Number);
      if (!box) box = { r1: rc[0], c1: rc[1], r2: rc[0], c2: rc[1] };
      box.r1 = Math.min(box.r1, rc[0]);
      box.c1 = Math.min(box.c1, rc[1]);
      box.r2 = Math.max(box.r2, rc[0]);
      box.c2 = Math.max(box.c2, rc[1]);
    }
    return box || { r1: 0, c1: 0, r2: 0, c2: 0 };
  };

  function findSheet(name) {
    return sheets.filter(function (s) {
      return s.name === name;
    })[0];
  }

  function makeContext() {
    var context = { pendingError: null };
    context.sync = function () {
      var error = context.pendingError;
      context.pendingError = null;
      return error ? Promise.reject(error) : Promise.resolve();
    };

    function makeRange(sheet, box) {
      var range = { load: function () {} };
      var label = sheet.name + "!" + colLetters(box.c1) + (box.r1 + 1) + ":" + colLetters(box.c2) + (box.r2 + 1);
      function forEachCell(value2d, fn) {
        for (var r = box.r1; r <= box.r2; r++)
          for (var c = box.c1; c <= box.c2; c++) {
            var v = Array.isArray(value2d) ? value2d[r - box.r1][c - box.c1] : value2d;
            fn(sheet.cell(r, c), v);
          }
      }
      function checkShape(value2d) {
        if (!Array.isArray(value2d)) return;
        if (value2d.length !== box.r2 - box.r1 + 1 || value2d[0].length !== box.c2 - box.c1 + 1)
          context.pendingError = { code: "InvalidArgument", message: "Kích thước mảng không khớp vùng " + label };
      }
      Object.defineProperty(range, "address", {
        get: function () {
          var a = colLetters(box.c1) + (box.r1 + 1);
          var b = colLetters(box.c2) + (box.r2 + 1);
          return quoteSheet(sheet.name) + "!" + (a === b ? a : a + ":" + b);
        },
      });
      Object.defineProperty(range, "rowIndex", {
        get: function () {
          return box.r1;
        },
      });
      Object.defineProperty(range, "rowCount", {
        get: function () {
          return box.r2 - box.r1 + 1;
        },
      });
      Object.defineProperty(range, "columnCount", {
        get: function () {
          return box.c2 - box.c1 + 1;
        },
      });
      Object.defineProperty(range, "columnIndex", {
        get: function () {
          return box.c1;
        },
      });
      Object.defineProperty(range, "values", {
        get: function () {
          var out = [];
          for (var r = box.r1; r <= box.r2; r++) {
            var row = [];
            for (var c = box.c1; c <= box.c2; c++) {
              var cell = sheet.peek(r, c);
              row.push(cell ? cell.value : "");
            }
            out.push(row);
          }
          return out;
        },
        set: function (v) {
          checkShape(v);
          var single = Array.isArray(v) && v.length === 1 && v[0].length === 1 ? v[0][0] : undefined;
          log.push({ op: "values", range: label, value: single });
          forEachCell(v, function (cell, value) {
            cell.formula = null;
            cell.value = typeof value === "string" && value.charAt(0) === "'" ? value.slice(1) : value;
          });
        },
      });
      Object.defineProperty(range, "formulas", {
        get: function () {
          var out = [];
          for (var r = box.r1; r <= box.r2; r++) {
            var row = [];
            for (var c = box.c1; c <= box.c2; c++) {
              var cell = sheet.peek(r, c);
              row.push(cell ? cell.formula || cell.value : "");
            }
            out.push(row);
          }
          return out;
        },
      });
      Object.defineProperty(range, "numberFormat", {
        set: function (v) {
          checkShape(v);
          log.push({ op: "numberFormat", range: label });
          forEachCell(v, function (cell, value) {
            cell.numberFormat = value;
          });
        },
      });
      var font = {};
      var format = { font: font };
      ["wrapText", "verticalAlignment", "horizontalAlignment", "columnWidth", "rowHeight"].forEach(function (p) {
        Object.defineProperty(format, p, {
          set: function (v) {
            if (p === "columnWidth" || p === "rowHeight") {
              if (input.apiMinor < 2) context.pendingError = { code: "ApiNotFound", message: p + " cần ExcelApi 1.2" };
            }
            log.push({ op: "format." + p, range: label, value: v });
            forEachCell(v, function (cell, value) {
              cell.format[p] = value;
            });
          },
        });
      });
      Object.defineProperty(font, "bold", {
        set: function (v) {
          log.push({ op: "font.bold", range: label, value: v });
        },
      });
      range.format = format;
      range.worksheet = sheetApi(sheet);
      range.getCell = function (r, c) {
        return makeRange(sheet, { r1: box.r1 + r, c1: box.c1 + c, r2: box.r1 + r, c2: box.c1 + c });
      };
      range.getIntersection = function (other) {
        var o = other.__box;
        var b = {
          r1: Math.max(box.r1, o.r1),
          c1: Math.max(box.c1, o.c1),
          r2: Math.min(box.r2, o.r2),
          c2: Math.min(box.c2, o.c2),
        };
        if (b.r1 > b.r2 || b.c1 > b.c2) {
          context.pendingError = { code: "ItemNotFound", message: "Không có giao" };
          return makeRange(sheet, box);
        }
        return makeRange(sheet, b);
      };
      range.__box = box;
      return range;
    }

    function sheetApi(sheet) {
      return {
        load: function () {},
        get name() {
          return sheet.name;
        },
        getRange: function (address) {
          try {
            return makeRange(sheet, parseAddress(address));
          } catch (error) {
            context.pendingError = error;
            return makeRange(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 });
          }
        },
        getUsedRange: function () {
          return makeRange(sheet, sheet.usedBox());
        },
        activate: function () {
          log.push({ op: "activate", sheet: sheet.name });
        },
      };
    }

    var worksheets = {
      load: function () {},
      get items() {
        return sheets.map(sheetApi);
      },
      getItem: function (name) {
        var found = sheets.filter(function (s) {
          return s.name.toLowerCase() === String(name).toLowerCase();
        })[0];
        if (!found) {
          context.pendingError = { code: "ItemNotFound", message: "Không có sheet " + name };
          return sheetApi(new Sheet(name));
        }
        return sheetApi(found);
      },
      add: function (name) {
        var sheet = new Sheet(name);
        sheets.push(sheet);
        log.push({ op: "addSheet", sheet: name });
        return sheetApi(sheet);
      },
    };

    context.workbook = {
      worksheets: worksheets,
      getSelectedRange: function () {
        return worksheets.getItem(selection.sheet).getRange(selection.address);
      },
    };
    return context;
  }

  input.sheets.forEach(function (spec) {
    var sheet = new Sheet(spec.name);
    spec.values.forEach(function (row, r) {
      row.forEach(function (value, c) {
        if (value !== "") sheet.cell(r, c).value = value;
      });
    });
    Object.keys(spec.formulas || {}).forEach(function (address) {
      var box = parseAddress(address);
      sheet.cell(box.r1, box.c1).formula = spec.formulas[address];
    });
    sheets.push(sheet);
  });

  window.__smoke = {
    log: log,
    dump: function (name) {
      var sheet = sheets.filter(function (s) {
        return s.name === name;
      })[0];
      if (!sheet) return null;
      var box = sheet.usedBox();
      var rows = [];
      for (var r = box.r1; r <= box.r2; r++) {
        var row = [];
        for (var c = box.c1; c <= box.c2; c++) {
          var cell = sheet.peek(r, c);
          row.push(cell ? { value: cell.value, numberFormat: cell.numberFormat } : null);
        }
        rows.push(row);
      }
      return rows;
    },
    select: function (sheetName, address) {
      selection = { sheet: sheetName, address: address };
    },
    setValue: function (sheetName, address, value) {
      var box = parseAddress(address);
      findSheet(sheetName).cell(box.r1, box.c1).value = value;
    },
    cell: function (sheetName, address) {
      var sheet = findSheet(sheetName);
      var box = parseAddress(address);
      var cell = sheet && sheet.peek(box.r1, box.c1);
      return cell ? { value: cell.value, formula: cell.formula } : null;
    },
    renameSheet: function (from, to) {
      findSheet(from).name = to;
    },
    settings: function () {
      return JSON.parse(JSON.stringify(settingsStore));
    },
    exportSheets: function () {
      return sheets.map(function (s) {
        var box = s.usedBox();
        var values = [];
        var formulas = {};
        for (var r = 0; r <= box.r2; r++) {
          var row = [];
          for (var c = 0; c <= box.c2; c++) {
            var cell = s.peek(r, c);
            row.push(cell ? cell.value : "");
            if (cell && cell.formula) formulas[colLetters(c) + (r + 1)] = cell.formula;
          }
          values.push(row);
        }
        return { name: s.name, values: values, formulas: formulas };
      });
    },
    deleteSheet: function (name) {
      sheets = sheets.filter(function (s) {
        return s.name !== name;
      });
    },
    openDialog: function () {
      var d = openDialog();
      return d ? { index: dialogs.indexOf(d), url: d.url, width: d.options.width, height: d.options.height } : null;
    },
    dialogClosed: function (index) {
      return dialogs[index].closed;
    },
    dialogCount: function () {
      return dialogs.length;
    },
    /** Tin từ hộp thoại; mặc định origin = origin của add-in (như Office gửi kèm, DialogOrigin 1.1). */
    deliver: function (message, origin) {
      var d = openDialog();
      var from = origin === undefined ? window.location.protocol + "//" + window.location.host : origin;
      if (d && d.handlers.dialogMessageReceived) d.handlers.dialogMessageReceived({ message: message, origin: from });
    },
    dismiss: function () {
      var d = openDialog();
      if (!d) return;
      d.closed = true;
      if (d.handlers.dialogEventReceived) d.handlers.dialogEventReceived({ error: 12006 });
    },
    sheetNames: function () {
      return sheets.map(function (s) {
        return s.name;
      });
    },
  };

  window.Excel = {
    run: function (callback) {
      return callback(makeContext());
    },
  };
  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { DialogMessageReceived: "dialogMessageReceived", DialogEventReceived: "dialogEventReceived" },
    onReady: function () {
      return Promise.resolve({ host: "Excel", platform: "PC" });
    },
    context: {
      document: {
        settings: {
          get: function (key) {
            return key in settingsStore ? JSON.parse(JSON.stringify(settingsStore[key])) : null;
          },
          set: function (key, value) {
            settingsStore[key] = JSON.parse(JSON.stringify(value));
          },
          saveAsync: function (callback) {
            log.push({ op: "settings.saveAsync" });
            if (callback) {
              setTimeout(function () {
                callback({ status: "succeeded" });
              }, 0);
            }
          },
        },
      },
      diagnostics: { platform: "PC", version: "16.0 (giả lập)" },
      ui: {
        displayDialogAsync: function (url, options, callback) {
          if (openDialog()) {
            setTimeout(function () {
              callback({ status: "failed", error: { code: 12007, message: "Đang có hộp thoại" } });
            }, 0);
            return;
          }
          var dialog = { url: url, options: options, closed: false, handlers: {} };
          dialog.api = {
            addEventHandler: function (type, handler) {
              dialog.handlers[type] = handler;
            },
            close: function () {
              dialog.closed = true;
            },
          };
          dialogs.push(dialog);
          log.push({ op: "dialog", url: url.split("#")[0] });
          setTimeout(function () {
            callback({ status: "succeeded", value: dialog.api });
          }, 0);
        },
        messageParent: function (message) {
          window.__toParent(message);
        },
      },
      requirements: {
        isSetSupported: function (name, version) {
          return name === "ExcelApi" && Number(String(version).split(".")[1]) <= input.apiMinor;
        },
      },
    },
  };
})();
