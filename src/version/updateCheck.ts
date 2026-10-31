import { type ReleaseInfo, parseReleaseInfo } from "../core/version/version";

const REQUEST_TIMEOUT_MS = 10_000;

/**
 * Tải `version.json` cùng thư mục với task pane (cùng origin → không vướng CORS).
 * Dùng XMLHttpRequest vì Internet Explorer 11 (Excel 2016/2019) không có `fetch`.
 * Trả về `undefined` khi lỗi mạng, hết thời gian chờ hoặc nội dung không hợp lệ.
 */
export function fetchReleaseInfo(): Promise<ReleaseInfo | undefined> {
  return new Promise((resolve) => {
    const request = new XMLHttpRequest();
    // Tham số chống cache: luôn lấy thông tin phát hành mới nhất.
    request.open("GET", `version.json?t=${String(Date.now())}`, true);
    request.timeout = REQUEST_TIMEOUT_MS;
    request.onload = () => {
      resolve(request.status === 200 ? parseReleaseInfo(request.responseText) : undefined);
    };
    request.onerror = () => {
      resolve(undefined);
    };
    request.ontimeout = () => {
      resolve(undefined);
    };
    request.send();
  });
}
