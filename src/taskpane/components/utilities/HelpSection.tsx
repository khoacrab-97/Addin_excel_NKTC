import { Link, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import { LINKS } from "../../../version/links";

const listStyle = { margin: "4px 0", paddingLeft: 18 } as const;

/** Hướng dẫn sử dụng và file mẫu (P3.1). */
export function HelpSection(): JSX.Element {
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Text variant="large">Hướng dẫn sử dụng</Text>

      <Text variant="mediumPlus">1. Xuất nhật ký theo ngày (tab Theo ngày)</Text>
      <ol style={listStyle}>
        <li>
          Bôi chọn bảng tiến độ trên sheet, gồm dòng tiêu đề (cột Nội dung công việc, Ngày bắt đầu, Ngày kết thúc).
        </li>
        <li>Nhấn &quot;Lấy vùng đang chọn&quot;; kiểm tra 3 cột được nhận diện, chỉnh lại nếu cần.</li>
        <li>Xem số công việc hợp lệ, các dòng lỗi và 5 ngày xem trước.</li>
        <li>
          Nhấn &quot;Xuất nhật ký theo ngày&quot; → add-in tạo sheet mới <strong>NK_TheoNgay</strong> (không ghi đè
          sheet có sẵn).
        </li>
      </ol>

      <Text variant="mediumPlus">2. Xuất theo form mẫu (tab Form mẫu)</Text>
      <ol style={listStyle}>
        <li>
          Trong sheet NK_TheoNgay, thêm các cột phụ khi cần: Thời tiết, Ban chỉ huy - CBKT, Số công nhân, Số thiết bị,
          Nội dung nghiệm thu, Sự cố, Kiến nghị.
        </li>
        <li>Chép sheet &quot;Mẫu NKTC&quot; (từ file biểu mẫu bên dưới) vào file nhật ký.</li>
        <li>Bôi chọn vùng dữ liệu NK_TheoNgay (gồm tiêu đề) → &quot;Lấy vùng đang chọn&quot;; kiểm tra cột Ngày.</li>
        <li>Với mỗi cột: chọn ô (hoặc ô merge) trên biểu mẫu → nhấn &quot;Ô đang chọn&quot; ở cột đó.</li>
        <li>Nhấn &quot;Xuất lên form&quot;, rồi chuyển ngày bằng « ‹ › » hoặc nhập ngày → &quot;Đi tới&quot;.</li>
        <li>Cấu hình gán cột được lưu trong file; lần sau mở lại dùng tiếp.</li>
      </ol>

      <Text variant="mediumPlus">Tải về</Text>
      <ul style={listStyle}>
        <li>
          <Link href={LINKS.formTemplate} target="_blank" rel="noopener noreferrer">
            Biểu mẫu &quot;Tình hình thi công hàng ngày&quot; (Mau_NKTC.xlsx)
          </Link>
        </li>
        <li>
          <Link href={LINKS.sampleSchedule} target="_blank" rel="noopener noreferrer">
            Danh sách công việc mẫu (Mau_TienDo.xlsx)
          </Link>
        </li>
        <li>
          <Link href={LINKS.formTargets} target="_blank" rel="noopener noreferrer">
            Bảng ô đích của biểu mẫu
          </Link>
        </li>
        <li>
          <Link href={LINKS.installer} target="_blank" rel="noopener noreferrer">
            Bộ cài cho máy khác (NhatKyThiCong_Setup.exe)
          </Link>
        </li>
        <li>
          <Link href={LINKS.installGuide} target="_blank" rel="noopener noreferrer">
            Hướng dẫn cài đặt, cập nhật, xử lý sự cố
          </Link>
        </li>
      </ul>
    </Stack>
  );
}
