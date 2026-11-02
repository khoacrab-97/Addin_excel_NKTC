import { APP_VERSION } from "../version/version";

/**
 * Trang chào mừng: chỉ hiện một lần khi bộ cài kích hoạt add-in (workbook NhatKyThiCong_BatDau.xlsx).
 * Mọi chức năng nằm trên tab Ribbon "Nhật ký thi công"; không còn khung làm việc bên phải.
 */
Office.onReady()
  .then(() => {
    const version = document.getElementById("version");
    if (version) version.textContent = APP_VERSION;
  })
  .catch(() => {
    // Trang tĩnh vẫn hiển thị nội dung khi Office.js không khởi động được.
  });
