# Contributing

## Quy trình

1. Tạo branch ngắn từ `main`: `feature/<ten>`, `fix/<ten>` hoặc `chore/<ten>`.
2. Chỉ thay đổi một mục đích trong mỗi commit.
3. Dùng Conventional Commits, ví dụ `fix: handle renamed Outlook folders`.
4. Chạy toàn bộ file `Tests/Test-*.ps1` bằng Windows PowerShell 5.1.
5. Mở pull request và mô tả tác động tới dữ liệu, Excel, Outlook hoặc installer.

## Yêu cầu trước khi mở PR

- Tất cả script PowerShell parse không lỗi.
- Tất cả test chạy thành công.
- Version trong `Launcher.ps1`, `Scripts/Build-Installer.ps1` và `Installer/AutoTools.iss` khớp nhau nếu chuẩn bị release.
- Không có workbook, output, log, profile trình duyệt, token hoặc đường dẫn user hard-code trong diff. Email thật chỉ được phép trong `Config/suppliers.csv` khi owner phê duyệt và repo vẫn private.
- README/changelog được cập nhật nếu thay đổi hành vi người dùng.

Không sửa dữ liệu vận hành hoặc refactor ngoài phạm vi của PR.
