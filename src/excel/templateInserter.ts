import { UserFacingError } from "./errors";

/** Tên sheet biểu mẫu trong `templates/Mau_NKTC.xlsx`. */
export const FORM_SHEET_NAME = "Mẫu NKTC";
/** Đường dẫn biểu mẫu (cùng thư mục với function file trên máy chủ phát hành). */
export const FORM_TEMPLATE_PATH = "templates/Mau_NKTC.xlsx";
const REQUEST_TIMEOUT_MS = 20_000;
const CHUNK = 0x8000;

/** Chèn sheet từ file mẫu cần ExcelApi 1.13 (Microsoft 365, Excel 2021 trở lên). */
export function canInsertTemplate(): boolean {
  return Office.context.requirements.isSetSupported("ExcelApi", "1.13");
}

/** Tải biểu mẫu dạng base64 (XMLHttpRequest — tương thích Internet Explorer 11). */
function fetchTemplateBase64(): Promise<string> {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("GET", `${FORM_TEMPLATE_PATH}?t=${String(Date.now())}`, true);
    request.responseType = "arraybuffer";
    request.timeout = REQUEST_TIMEOUT_MS;
    const fail = (): void => {
      reject(new UserFacingError("Không tải được biểu mẫu từ máy chủ phát hành. Vui lòng kiểm tra Internet."));
    };
    request.onload = () => {
      if (request.status !== 200 || !(request.response instanceof ArrayBuffer)) {
        fail();
        return;
      }
      const bytes = new Uint8Array(request.response);
      let binary = "";
      for (let offset = 0; offset < bytes.length; offset += CHUNK) {
        binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(offset, offset + CHUNK)));
      }
      resolve(window.btoa(binary));
    };
    request.onerror = fail;
    request.ontimeout = fail;
    request.send();
  });
}

/**
 * Chèn sheet "Mẫu NKTC" vào cuối workbook (F2). Không ghi đè: nếu đã có sheet cùng tên thì báo lỗi.
 * @returns Tên sheet đã chèn.
 */
export async function insertFormTemplate(): Promise<string> {
  const base64 = await fetchTemplateBase64();
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    if (worksheets.items.some((sheet) => sheet.name.toLowerCase() === FORM_SHEET_NAME.toLowerCase())) {
      throw new UserFacingError(
        `File đã có sheet "${FORM_SHEET_NAME}". Add-in không ghi đè — đổi tên hoặc xóa sheet đó nếu muốn chèn mẫu mới.`,
      );
    }
    const ids = context.workbook.insertWorksheetsFromBase64(base64, {
      sheetNamesToInsert: [FORM_SHEET_NAME],
      positionType: "End",
    });
    await context.sync();
    const id = ids.value[0];
    if (id === undefined) throw new UserFacingError("Không chèn được biểu mẫu.");
    const sheet = worksheets.getItem(id);
    sheet.load("name");
    sheet.activate();
    await context.sync();
    return sheet.name;
  });
}
