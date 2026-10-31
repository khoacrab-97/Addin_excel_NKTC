/** Thông tin môi trường Office/Excel nơi add-in đang chạy. */
export interface HostEnvironment {
  platform: string;
  officeVersion: string;
  /** Phiên bản ExcelApi cao nhất được hỗ trợ, ví dụ "1.1" (Excel 2016) hoặc "1.17". */
  maxExcelApi: string;
}

/** Giới hạn trên khi dò requirement set ExcelApi 1.x. */
const EXCEL_API_PROBE_LIMIT = 30;

/** Dò phiên bản ExcelApi cao nhất được hỗ trợ bằng `isSetSupported`. */
export function detectMaxExcelApi(): string {
  let maxMinor = 0;
  for (let minor = 1; minor <= EXCEL_API_PROBE_LIMIT; minor++) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", `1.${minor}`)) break;
    maxMinor = minor;
  }
  return maxMinor === 0 ? "Không hỗ trợ" : `1.${maxMinor}`;
}

/** Đọc thông tin môi trường. Gọi sau khi `Office.onReady` hoàn tất. */
export function getHostEnvironment(): HostEnvironment {
  // `diagnostics` có thể không có trên bản Office cũ → kiểm tra trước khi đọc.
  const diagnostics: Office.ContextInformation | undefined = Office.context.diagnostics;
  return {
    platform: diagnostics ? String(diagnostics.platform) : "Không xác định",
    officeVersion: diagnostics ? diagnostics.version : "Không xác định",
    maxExcelApi: detectMaxExcelApi(),
  };
}
