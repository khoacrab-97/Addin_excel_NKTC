/** Lỗi đã có thông điệp tiếng Việt để hiển thị trực tiếp cho người dùng. */
export class UserFacingError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UserFacingError";
    // Giữ đúng prototype khi transpile xuống ES5 (Internet Explorer 11).
    Object.setPrototypeOf(this, UserFacingError.prototype);
  }
}

/** Chuyển lỗi bất kỳ (Office.js, JavaScript) thành thông điệp cho người dùng. */
export function toUserMessage(error: unknown): string {
  if (error instanceof UserFacingError) return error.message;
  const code = typeof error === "object" && error !== null && "code" in error ? String(error.code) : "";
  switch (code) {
    case "ItemNotFound":
      return "Không tìm thấy sheet hoặc vùng dữ liệu. Sheet có thể đã bị xóa hoặc đổi tên.";
    case "InvalidArgument":
    case "InvalidReference":
      return "Địa chỉ vùng dữ liệu không hợp lệ.";
    case "AccessDenied":
    case "InvalidOperation":
      return "Excel không cho phép thao tác này (sheet có thể đang được bảo vệ hoặc đang ở chế độ chỉnh sửa ô).";
    default:
      return `Lỗi không xác định: ${error instanceof Error ? error.message : String(error)}`;
  }
}
