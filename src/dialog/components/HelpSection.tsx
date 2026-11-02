import { Link, Stack, Text } from "@fluentui/react";
import type { JSX } from "react";
import { LINKS } from "../../version/links";

const listStyle = { margin: "4px 0", paddingLeft: 18 } as const;

/** Hướng dẫn sử dụng và file mẫu (P3.1). */
export function HelpSection(): JSX.Element {
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Text>
        Mọi chức năng nằm trên tab <strong>Nhật ký thi công</strong> của thanh Ribbon.
      </Text>

      <Text variant="mediumPlus">1. Xuất nhật ký theo ngày (nhóm Xuất nhật ký)</Text>
      <ol style={listStyle}>
        <li>
          Bôi chọn bảng tiến độ (gồm dòng tiêu đề: Nội dung công việc, Ngày bắt đầu, Ngày kết thúc) — hoặc chỉ cần đặt
          con trỏ trong sheet tiến độ.
        </li>
        <li>
          Bấm &quot;Xuất theo ngày&quot;; kiểm tra 3 cột được nhận diện, chỉnh lại nếu cần → &quot;Xuất nhật ký&quot;.
        </li>
        <li>
          Add-in tạo sheet mới <strong>NK_TheoNgay</strong> (không ghi đè sheet có sẵn) và báo kết quả, các dòng lỗi.
        </li>
      </ol>

      <Text variant="mediumPlus">2. Xuất theo form mẫu (nhóm Form mẫu và Chuyển ngày)</Text>
      <ol style={listStyle}>
        <li>
          Trong sheet NK_TheoNgay, thêm các cột phụ khi cần: Thời tiết, Ban chỉ huy - CBKT, Số công nhân, Số thiết bị,
          Nội dung nghiệm thu, Sự cố, Kiến nghị.
        </li>
        <li>Chép sheet &quot;Mẫu NKTC&quot; (từ file biểu mẫu bên dưới) vào file nhật ký.</li>
        <li>Đặt con trỏ trong sheet NK_TheoNgay → bấm &quot;Vùng nguồn&quot;.</li>
        <li>
          Trên biểu mẫu: chọn một ô (hoặc ô merge) → bấm &quot;Gán cột&quot; → chọn cột ghi vào ô đó. Lặp lại cho từng
          ô.
        </li>
        <li>
          Bấm &quot;Ngày đầu&quot;, &quot;Ngày trước&quot;, &quot;Ngày sau&quot;, &quot;Ngày cuối&quot; hoặc &quot;Đến
          ngày…&quot; để ghi dữ liệu của ngày đó lên biểu mẫu. Lần ghi đầu, nếu ô đích đang có dữ liệu/công thức, add-in
          hỏi xác nhận.
        </li>
        <li>&quot;Danh sách gán&quot;: xem, bỏ gán ô, đổi cột Ngày. Cấu hình được lưu trong file Excel.</li>
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
