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

if (!fs.existsSync(manifest)) {
  console.error("Không thấy dist/manifest.xml — hãy chạy `npm run build` trước.");
  process.exit(1);
}
if (fs.readFileSync(manifest, "utf8").includes("localhost")) {
  console.error("dist/manifest.xml đang trỏ localhost — bộ cài chỉ đóng gói bản production (`npm run build`).");
  process.exit(1);
}

execFileSync(
  "makensis",
  ["-V2", `-DVERSION=${pkg.version}`, `-DMANIFEST=${manifest}`, `-DOUTFILE=${output}`, "NhatKyThiCong.nsi"],
  { cwd: path.join(root, "installer"), stdio: "inherit" },
);
console.log(`Đã tạo ${path.relative(root, output)} (${fs.statSync(output).size} bytes), phiên bản ${pkg.version}`);
