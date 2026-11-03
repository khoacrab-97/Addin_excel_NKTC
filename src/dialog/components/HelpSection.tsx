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

      <Text variant="mediumPlus">0. Tạo bảng nhập tiến độ (nhóm Tiến độ)</Text>
      <ol style={listStyle}>
        <li>
          Bấm &quot;Tạo bảng&quot; → chọn sheet mới hoặc sheet hiện tại (nhập ô bắt đầu), số dòng (mặc định 10) →
          &quot;Tạo bảng&quot;.
        </li>
        <li>
          Điền Nội dung công việc, ngày Bắt đầu, ngày Kết thúc (dd/mm/yyyy). Gõ ngay dưới dòng cuối, bảng tự mở rộng.
        </li>
      </ol>

      <Text variant="mediumPlus">1. Xuất nhật ký theo ngày (nhóm Tiến độ)</Text>
      <ol style={listStyle}>
        <li>
          Bôi chọn bảng tiến độ (gồm dòng tiêu đề: Nội dung công việc, Ngày bắt đầu, Ngày kết thúc) — hoặc chỉ cần đặt
          con trỏ trong sheet tiến độ.
        </li>
        <li>
          Bấm &quot;Xuất theo ngày&quot;; kiểm tra Vùng dữ liệu (chọn lại: bấm ⬆ → hộp thoại thu gọn thành thanh chọn →
          kéo chọn vùng trên sheet, địa chỉ hiện trong thanh → bấm ⬇ hoặc Enter) và 3 cột được nhận diện, chỉnh lại nếu
          cần → &quot;Xuất nhật ký&quot;.
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
        <li>
          Bấm &quot;Chèn form mẫu&quot; để thêm sheet &quot;Mẫu NKTC&quot; (Excel cũ: tải biểu mẫu bên dưới rồi chép
          sheet).
        </li>
        <li>
          Bấm &quot;Xuất vào form&quot;: kiểm tra Vùng nguồn (tự nhận diện; chọn lại bằng nút ⬆) — dòng đầu là tiêu đề,
          không được có ô gộp (merge).
        </li>
        <li>
          Mỗi dòng gán: chọn cột nguồn (trái); bên phải bấm ⬆ → bấm chọn ô trên biểu mẫu → bấm ⬇ (hoặc Enter) để quay
          lại. &quot;+ Thêm&quot; để gán tiếp; chọn dòng thừa → Delete (hoặc Xóa dòng).
        </li>
        <li>Bấm &quot;Gán&quot;: biểu mẫu nhận dữ liệu ngày đầu tiên. Cấu hình được lưu trong file Excel.</li>
        <li>
          Sau khi gán: dùng các nút &quot;Ngày đầu&quot;, &quot;Ngày trước&quot;, &quot;Ngày sau&quot;, &quot;Ngày
          cuối&quot; (nhóm Chuyển ngày). Chưa gán thì các nút này chỉ hiện hướng dẫn. Gán cột Ngày vào dòng &quot;Ngày …
          tháng … năm …&quot; của biểu mẫu: add-in tự điền ngày, tháng, năm sau từng chữ.
        </li>
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
