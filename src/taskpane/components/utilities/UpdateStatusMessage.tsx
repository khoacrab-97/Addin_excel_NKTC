import { Link, MessageBar, MessageBarType } from "@fluentui/react";
import type { JSX } from "react";
import type { UpdateStatus } from "../../../core/version/version";
import { LINKS } from "../../../version/links";

/** Hướng dẫn cập nhật khi có phiên bản mới (P3.3). */
export function UpdateSteps(): JSX.Element {
  return (
    <ol style={{ margin: "4px 0", paddingLeft: 18 }}>
      <li>Đóng khung add-in rồi mở lại (nút Nhật ký thi công trên tab Home).</li>
      <li>
        Nếu vẫn là bản cũ: đóng Excel, xóa nội dung thư mục <code>%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\</code>, mở
        lại Excel.
      </li>
    </ol>
  );
}

/** Kết quả kiểm tra cập nhật (P3.2). */
export function UpdateStatusMessage({ status }: { status: UpdateStatus }): JSX.Element {
  switch (status.kind) {
    case "latest":
      return (
        <MessageBar messageBarType={MessageBarType.success}>
          Bạn đang dùng phiên bản mới nhất ({status.current}).
        </MessageBar>
      );
    case "available":
      return (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <div>
            Đã có phiên bản <strong>{status.latest}</strong> (bạn đang dùng {status.current}). Cách cập nhật:
          </div>
          <UpdateSteps />
          <Link href={LINKS.updateGuide} target="_blank" rel="noopener noreferrer">
            Xem hướng dẫn cập nhật chi tiết
          </Link>
        </MessageBar>
      );
    case "newer-than-release":
      return (
        <MessageBar messageBarType={MessageBarType.info}>
          Bạn đang dùng bản phát triển {status.current}, mới hơn bản phát hành {status.latest}.
        </MessageBar>
      );
    case "unknown":
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          Không kiểm tra được phiên bản mới. Vui lòng kiểm tra kết nối Internet rồi thử lại.
        </MessageBar>
      );
  }
}
