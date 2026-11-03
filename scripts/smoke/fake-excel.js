/*
 * Excel/Office giả lập (chạy trong trình duyệt) cho smoke test. KHÔNG dùng trong add-in.
 * Mô phỏng một phần Office.js ExcelApi 1.1 đủ cho các luồng của add-in:
 * workbook trong bộ nhớ, getRange/getUsedRange/getIntersection/getSelectedRange,
 * ghi values/numberFormat/format, lỗi ItemNotFound được trả về lúc context.sync().
 * Đầu vào: window.__SMOKE_INPUT = {
 *   sheets: [{ name, values, formulas?: { "A26": "=..." } }], selection: { sheet, address }, apiMinor, settings? }.
 * Ghi chuỗi bắt đầu bằng "'" → lưu phần sau dấu "'" (giống Excel: dấu "'" là ký tự đánh dấu văn bản).
 * ExcelApi 1.13 (khi apiMinor ≥ 13): getMergedAreasOrNullObject (ô gộp khai báo trong spec.merged, ví dụ ["A1:B1"]),
 * insertWorksheetsFromBase64 (chèn sheet từ input.templateSheets; nội dung base64 không được đọc).
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
  /** Thời gian Excel cần để đóng hẳn một hộp thoại (mô phỏng). */
  var CLOSE_DELAY_MS = input.dialogCloseDelayMs === undefined ? 300 : input.dialogCloseDelayMs;
  var closingUntil = 0;
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

  var activeSheetName = input.selection.sheet;
  /** Hàm nghe DocumentSelectionChanged (Office.context.document.addHandlerAsync). */
  var selectionHandlers = [];
  var tableCount = 0;
  var shapeCount = 0;
  var COL_PT = 48;
  var ROW_PT = 15;
  function Sheet(name) {
    this.name = name;
    this.cells = {};
    /** Excel Table giả lập: { name, box, style } — báo lỗi khi chồng lên bảng có sẵn (giống Excel). */
    this.tables = [];
    this.merged = [];
    /** Shape trên sheet: { id, type, name, left, top, width, height, fill, line }. */
    this.shapes = [];
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
      // Vị trí (point) giả lập: cột rộng 48pt, dòng cao 15pt.
      [
        [
          "left",
          function () {
            return box.c1 * COL_PT;
          },
        ],
        [
          "top",
          function () {
            return box.r1 * ROW_PT;
          },
        ],
        [
          "width",
          function () {
            return (box.c2 - box.c1 + 1) * COL_PT;
          },
        ],
        [
          "height",
          function () {
            return (box.r2 - box.r1 + 1) * ROW_PT;
          },
        ],
      ].forEach(function (pair) {
        Object.defineProperty(range, pair[0], { get: pair[1] });
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
      range.getRow = function (r) {
        return makeRange(sheet, { r1: box.r1 + r, c1: box.c1, r2: box.r1 + r, c2: box.c2 });
      };
      range.getMergedAreasOrNullObject = function () {
        if (input.apiMinor < 13) context.pendingError = { code: "ApiNotFound", message: "cần ExcelApi 1.13" };
        var hits = sheet.merged.filter(function (m) {
          return !(box.r2 < m.r1 || box.r1 > m.r2 || box.c2 < m.c1 || box.c1 > m.c2);
        });
        var address = hits
          .map(function (m) {
            return quoteSheet(sheet.name) + "!" + colLetters(m.c1) + (m.r1 + 1) + ":" + colLetters(m.c2) + (m.r2 + 1);
          })
          .join(",");
        return { load: function () {}, isNullObject: hits.length === 0, address: address };
      };
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
      range.select = function () {
        selection = { sheet: sheet.name, address: label.split("!")[1] };
        activeSheetName = sheet.name;
        log.push({ op: "select", range: label });
      };
      range.__box = box;
      return range;
    }

    /** Shape giả lập (ExcelApi 1.9): chỉ phần add-in dùng cho viền nháy. */
    function shapesApi(sheet) {
      function wrap(shape) {
        var api = {
          load: function () {},
          isNullObject: false,
          get id() {
            return shape.id;
          },
          delete: function () {
            sheet.shapes = sheet.shapes.filter(function (item) {
              return item !== shape;
            });
            log.push({ op: "shape.delete", sheet: sheet.name });
          },
          fill: {
            clear: function () {
              shape.fill = "none";
            },
          },
          lineFormat: {},
        };
        ["name", "left", "top", "width", "height"].forEach(function (p) {
          Object.defineProperty(api, p, {
            get: function () {
              return shape[p];
            },
            set: function (v) {
              shape[p] = v;
            },
          });
        });
        ["color", "weight", "dashStyle"].forEach(function (p) {
          Object.defineProperty(api.lineFormat, p, {
            set: function (v) {
              shape.line[p] = v;
              if (p === "color") log.push({ op: "shape.color", sheet: sheet.name, value: v });
            },
          });
        });
        return api;
      }
      return {
        load: function () {},
        get items() {
          return sheet.shapes.map(wrap);
        },
        addGeometricShape: function (type) {
          if (input.apiMinor < 9) context.pendingError = { code: "ApiNotFound", message: "cần ExcelApi 1.9" };
          var shape = { id: "shape" + ++shapeCount, type: type, line: {}, fill: "solid" };
          sheet.shapes.push(shape);
          log.push({ op: "shape.add", sheet: sheet.name, type: type });
          return wrap(shape);
        },
        getItemOrNullObject: function (id) {
          var found = sheet.shapes.filter(function (item) {
            return item.id === id;
          })[0];
          return found ? wrap(found) : { load: function () {}, isNullObject: true };
        },
      };
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
          activeSheetName = sheet.name;
          log.push({ op: "activate", sheet: sheet.name });
        },
        shapes: shapesApi(sheet),
        tables: {
          add: function (address, hasHeaders) {
            var box = parseAddress(address);
            var overlap = sheet.tables.some(function (t) {
              return !(box.r2 < t.box.r1 || box.r1 > t.box.r2 || box.c2 < t.box.c1 || box.c1 > t.box.c2);
            });
            if (overlap) context.pendingError = { code: "InvalidOperation", message: "Chồng lên bảng có sẵn" };
            var table = { name: "Table" + ++tableCount, box: box, hasHeaders: hasHeaders, style: "" };
            if (!overlap) sheet.tables.push(table);
            log.push({ op: "tables.add", sheet: sheet.name, address: address, hasHeaders: hasHeaders });
            var api = {};
            Object.defineProperty(api, "style", {
              set: function (v) {
                table.style = v;
              },
            });
            return api;
          },
        },
      };
    }

    var worksheets = {
      load: function () {},
      getActiveWorksheet: function () {
        return worksheets.getItem(activeSheetName);
      },
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
      insertWorksheetsFromBase64: function (base64, options) {
        if (input.apiMinor < 13) context.pendingError = { code: "ApiNotFound", message: "cần ExcelApi 1.13" };
        var names = (options && options.sheetNamesToInsert) || [];
        var ids = [];
        (input.templateSheets || []).forEach(function (spec) {
          if (names.length && names.indexOf(spec.name) === -1) return;
          sheets.push(buildSheet(spec));
          ids.push(spec.name);
        });
        log.push({ op: "insertWorksheets", names: ids, base64Length: base64.length });
        return { value: ids };
      },
      getSelectedRange: function () {
        return worksheets.getItem(selection.sheet).getRange(selection.address);
      },
    };
    return context;
  }

  function buildSheet(spec) {
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
    (spec.merged || []).forEach(function (address) {
      sheet.merged.push(parseAddress(address));
    });
    return sheet;
  }

  input.sheets.forEach(function (spec) {
    sheets.push(buildSheet(spec));
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
    /** Người dùng chọn trên sheet → Excel phát sự kiện DocumentSelectionChanged. */
    select: function (sheetName, address) {
      selection = { sheet: sheetName, address: address };
      activeSheetName = sheetName;
      selectionHandlers.slice().forEach(function (handler) {
        setTimeout(function () {
          handler({ type: "documentSelectionChanged" });
        }, 0);
      });
    },
    selectionHandlerCount: function () {
      return selectionHandlers.length;
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
    /** Chèn cột trống tại `columnIndex` (0-based) như Insert Column của Excel: các ô từ cột đó dịch sang phải. */
    insertColumn: function (sheetName, columnIndex) {
      var sheet = findSheet(sheetName);
      var moved = {};
      Object.keys(sheet.cells).forEach(function (key) {
        var rc = key.split(",").map(Number);
        moved[rc[0] + "," + (rc[1] >= columnIndex ? rc[1] + 1 : rc[1])] = sheet.cells[key];
      });
      sheet.cells = moved;
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
      closingUntil = Date.now() + CLOSE_DELAY_MS;
      if (d.handlers.dialogEventReceived) d.handlers.dialogEventReceived({ error: 12006 });
    },
    tables: function (name) {
      return findSheet(name).tables.map(function (t) {
        return {
          name: t.name,
          address: colLetters(t.box.c1) + (t.box.r1 + 1) + ":" + colLetters(t.box.c2) + (t.box.r2 + 1),
          style: t.style,
        };
      });
    },
    activeSheet: function () {
      return activeSheetName;
    },
    selection: function () {
      return selection;
    },
    /** Shape đang có trên sheet (viền nháy): tên, vị trí, kiểu viền. */
    shapes: function (name) {
      return findSheet(name).shapes.map(function (shape) {
        return {
          name: shape.name,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
          fill: shape.fill,
          dash: shape.line.dashStyle,
          color: shape.line.color,
        };
      });
    },
    geometry: { col: COL_PT, row: ROW_PT },
    /** Giả lập viền nháy còn sót (Excel bị đóng khi thanh chọn đang mở). */
    addShape: function (sheetName, name) {
      findSheet(sheetName).shapes.push({ id: "left" + ++shapeCount, name: name, line: {}, fill: "none" });
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
    EventType: {
      DialogMessageReceived: "dialogMessageReceived",
      DialogEventReceived: "dialogEventReceived",
      DialogParentMessageReceived: "dialogParentMessageReceived",
      DocumentSelectionChanged: "documentSelectionChanged",
    },
    onReady: function () {
      return Promise.resolve({ host: "Excel", platform: "PC" });
    },
    context: {
      document: {
        addHandlerAsync: function (type, handler, callback) {
          if (type === "documentSelectionChanged") selectionHandlers.push(handler);
          log.push({ op: "addHandler", type: type });
          setTimeout(function () {
            if (callback) callback({ status: type === "documentSelectionChanged" ? "succeeded" : "failed" });
          }, 0);
        },
        removeHandlerAsync: function (type, options) {
          var handler = options && options.handler;
          selectionHandlers = selectionHandlers.filter(function (item) {
            return item !== handler;
          });
          log.push({ op: "removeHandler", type: type });
        },
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
          // Giống Excel thật: hộp thoại vừa đóng còn "đang đóng" một lúc → mở hộp thoại mới báo 12007.
          if (openDialog() || Date.now() < closingUntil) {
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
            messageChild: function (message) {
              log.push({ op: "messageChild" });
              window.__toChild(message);
            },
            close: function () {
              if (!dialog.closed) closingUntil = Date.now() + CLOSE_DELAY_MS;
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
        /** Trang hộp thoại: nhận tin từ add-in (DialogApi 1.2). */
        addHandlerAsync: function (type, handler) {
          window.__parentHandler = handler;
        },
      },
      requirements: {
        isSetSupported: function (name, version) {
          var minor = Number(String(version).split(".")[1]);
          if (name === "DialogApi") return minor <= (input.dialogApiMinor === undefined ? 2 : input.dialogApiMinor);
          return name === "ExcelApi" && minor <= input.apiMinor;
        },
      },
    },
  };
})();
