# Security Policy

## Supported version

| Version | Supported |
| --- | --- |
| 1.2.x | Yes |
| 1.1.x and older | No |

## Báo cáo lỗ hổng

Không mở issue có chứa thông tin nhạy cảm. Hãy dùng **Security → Advisories → Report a vulnerability** nếu tính năng này được bật, hoặc liên hệ maintainer qua hồ sơ [hugo-lcqh](https://github.com/hugo-lcqh).

Khi báo cáo, chỉ cung cấp:

- Version Auto Tools và version Windows PowerShell.
- Tool/script bị ảnh hưởng và các bước tái hiện tối thiểu.
- Tác động dự kiến.
- Log đã che đường dẫn user, email, supplier, dữ liệu PO/CD và nội dung workbook.

Không gửi workbook sản xuất, attachment Outlook, `Config/suppliers.csv` hoặc dữ liệu trong `Input`, `Output`, `Logs` và `Temp`.

## Phạm vi dữ liệu

Auto Tools xử lý dữ liệu nghiệp vụ cục bộ và dùng Excel/Outlook COM. Repo phải được giữ private nếu còn chứa danh sách liên hệ supplier thật. Credential, token, mật khẩu và file dữ liệu vận hành không được commit.
