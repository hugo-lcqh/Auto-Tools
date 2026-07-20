# Auto Tools

Bộ công cụ PowerShell dành cho Windows, giúp tự động hóa các nghiệp vụ Material Control, xử lý dữ liệu Excel và Outlook.

**Current version:** 1.1.0  
**Maintainer:** Hugo Le Chi Quoc Hung

## Chức năng

- Tạo và kiểm tra Job Shortage.
- Kiểm tra thiếu hàng dựa trên Supplier Commitment.
- Kiểm tra MR Pull-in.
- Tải Supplier Commitment và Vendor Stock.
- Tạo file import CD và PO.
- Tạo email nháp Pre-CD trong Outlook.
- Launcher tập trung để chạy và quản lý các công cụ.
- Tạo file EXE và bộ cài đặt Windows.

## Yêu cầu

- Windows.
- Windows PowerShell 5.1 trở lên.
- Microsoft Excel Desktop cho các chức năng đọc và ghi Excel.
- Microsoft Outlook Desktop cho các chức năng liên quan đến email.
- Inno Setup 6 nếu cần build bộ cài đặt.

## Cài đặt

### Sử dụng bộ cài

Tải file `AutoTools-Setup-1.1.0.exe` từ phần **Releases**, sau đó chạy bộ cài.

### Chạy từ mã nguồn

```powershell
git clone https://github.com/moi0329/Auto-Tools.git
Set-Location Auto-Tools
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Launcher.ps1"
Cấu hình
Cấu hình đường dẫn dữ liệu tại:
Config/autotools-paths.json
Thông tin supplier được quản lý tại:
Config/suppliers.csv
Không commit dữ liệu công việc thực tế trong các thư mục Input, Output, Logs, Temp và dist.
Kiểm tra installer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Tests\Test-Installer.ps1"
Cấu trúc chính
Auto-Tools/
├── Config/
├── Input/
├── Installer/
├── Scripts/
├── Tests/
├── Launcher.ps1
└── launcher-config.json
Tài liệu
Xem hướng dẫn đầy đủ tại [HUONG_DAN_SETUP_AUTO_TOOLS.txt](./HUONG_DAN_SETUP_AUTO_TOOLS.txt)
