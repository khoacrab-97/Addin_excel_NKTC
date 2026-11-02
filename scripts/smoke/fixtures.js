/*
 * Đọc dữ liệu test cho smoke test. Chỉ nhận TÊN file .json trong test/fixtures/
 * (bỏ phần thư mục của tham số dòng lệnh) → không đọc được file ngoài thư mục này.
 */
const fs = require("fs");
const path = require("path");

const FIXTURES_DIR = path.resolve(__dirname, "..", "..", "test", "fixtures");

/** Đường dẫn file trong test/fixtures từ tham số (có thể kèm thư mục, ví dụ `test/fixtures/demo_f2.json`). */
function fixturePath(arg) {
  const name = path.basename(String(arg));
  if (!/^[\w.-]+\.json$/.test(name)) throw new Error("Tên file dữ liệu test không hợp lệ (cần *.json).");
  return path.join(FIXTURES_DIR, name);
}

function readFixture(arg) {
  return JSON.parse(fs.readFileSync(fixturePath(arg), "utf8"));
}

module.exports = { FIXTURES_DIR, fixturePath, readFixture };
