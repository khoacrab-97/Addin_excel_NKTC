/*
 * Đóng gói bộ cài Windows (NSIS) từ manifest production trong dist/.
 * Yêu cầu: đã chạy `npm run build` (production) và có `makensis` (NSIS 3; Ubuntu: apt-get install nsis).
 * Kết quả: dist/NhatKyThiCong_Setup.exe (đường dẫn cố định, luôn là bản mới nhất trên GitHub Pages).
 */
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const pkg = require("../package.json");

const root = path.join(__dirname, "..");
const manifest = path.join(root, "dist", "manifest.xml");
const output = path.join(root, "dist", "NhatKyThiCong_Setup.exe");

// Gọi makensis bằng đường dẫn tuyệt đối trong thư mục cài đặt cố định — không tìm theo PATH
// (tránh chạy nhầm chương trình cùng tên đặt trong thư mục ghi được; Sonar S4036 / CWE-426).
const MAKENSIS_CANDIDATES = [
  "/usr/bin/makensis",
  "/usr/local/bin/makensis",
  "/opt/homebrew/bin/makensis",
  "C:\\Program Files (x86)\\NSIS\\makensis.exe",
  "C:\\Program Files\\NSIS\\makensis.exe",
];
const makensis = MAKENSIS_CANDIDATES.find((candidate) => fs.existsSync(candidate));
if (makensis === undefined) {
  console.error(`Không thấy makensis (NSIS 3). Đã tìm: ${MAKENSIS_CANDIDATES.join(", ")}`);
  process.exit(1);
}

if (!fs.existsSync(manifest)) {
  console.error("Không thấy dist/manifest.xml — hãy chạy `npm run build` trước.");
  process.exit(1);
}
if (fs.readFileSync(manifest, "utf8").includes("localhost")) {
  console.error("dist/manifest.xml đang trỏ localhost — bộ cài chỉ đóng gói bản production (`npm run build`).");
  process.exit(1);
}

execFileSync(
  makensis,
  ["-V2", `-DVERSION=${pkg.version}`, `-DMANIFEST=${manifest}`, `-DOUTFILE=${output}`, "NhatKyThiCong.nsi"],
  { cwd: path.join(root, "installer"), stdio: "inherit" },
);
console.log(`Đã tạo ${path.relative(root, output)} (${fs.statSync(output).size} bytes), phiên bản ${pkg.version}`);
