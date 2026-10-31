# Nguồn skill

- Repo: https://github.com/cloudflare/security-audit-skill
- Thư mục: `skills/security-audit/` — commit `c1c8a8c` (31/10/2026), giấy phép MIT (xem `LICENSE`).
- Nội dung giữ nguyên, không chỉnh sửa. Đã rà soát trước khi đưa vào repo: hướng dẫn chỉ mang tính phòng thủ, validator chỉ đọc file JSON được chỉ định (không mạng, không ghi file, không chạy lệnh).
- Cập nhật: tải lại thư mục trên từ repo gốc, rà soát lại, rồi thay thế toàn bộ thư mục này.

Cách dùng trong dự án: yêu cầu Claude Code "security audit this codebase" (kiểm tra toàn diện) hoặc hỏi câu hỏi bảo mật cụ thể (chế độ hướng dẫn). Báo cáo mặc định ghi ra `~/security-audit-skill/<repo>/run-<N>` — ngoài repo.
