# Auto Tools

[![Version](https://img.shields.io/badge/version-1.2.0-2563eb)](https://github.com/hugo-lcqh/Auto-Tools/releases/latest)
[![PowerShell](https://img.shields.io/badge/Windows%20PowerShell-5.1-012456?logo=powershell)](https://learn.microsoft.com/powershell/)
[![PowerShell CI](https://github.com/hugo-lcqh/Auto-Tools/actions/workflows/ci.yml/badge.svg)](https://github.com/hugo-lcqh/Auto-Tools/actions/workflows/ci.yml)

Bộ công cụ Windows PowerShell dành cho Material Control, tự động hóa các quy trình Excel và Outlook như kiểm tra shortage, Supplier Commitment, MR Pull-in, Pre-CD, PO và CD.

> Phiên bản ổn định hiện tại: **v1.2.0** · Duy trì bởi [Hugo Le Chi Quoc Hung](https://github.com/hugo-lcqh)

## Tính năng chính

| Nhóm | Chức năng |
| --- | --- |
| Job Shortage | Tạo dữ liệu ngày mới, archive dữ liệu cũ và kiểm tra shortage/CD |
| Supplier Commitment | Tải attachment từ Outlook, phát hiện supplier còn thiếu và tạo email nhắc dạng draft |
| MR Pull-in | Phân tích và kiểm tra nhu cầu pull-in |
| Vendor Stock | Tải dữ liệu tồn kho supplier từ Outlook |
| Pre-CD | Tạo file import CD và email draft trong Outlook |
| PO | Tạo file PO theo từng supplier hoặc gộp theo site TC5/TN5 |
| Launcher | Giao diện tập trung, hỗ trợ DPI scaling, accessibility và kiểm tra trạng thái tool |
| Distribution | Build file EXE và bộ cài Windows bằng Inno Setup |

## Yêu cầu hệ thống

- Windows 10/11.
- Windows PowerShell 5.1.
- Microsoft Excel Desktop cho các chức năng đọc/ghi workbook.
- Microsoft Outlook Desktop cho các chức năng email.
- Inno Setup 6 chỉ cần khi build bộ cài.

## Cài đặt

### Từ Release

Tải `AutoTools-Setup-1.2.0.exe` tại [GitHub Releases](https://github.com/hugo-lcqh/Auto-Tools/releases/latest), sau đó chạy bộ cài. Ứng dụng được cài theo từng user tại `%LOCALAPPDATA%\Programs\Auto Tools` và không yêu cầu quyền Administrator.

### Từ mã nguồn

```powershell
git clone https://github.com/hugo-lcqh/Auto-Tools.git
Set-Location Auto-Tools
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Launcher.ps1"
```

## Cấu hình

- Đường dẫn dữ liệu: [`Config/autotools-paths.json`](Config/autotools-paths.json)
- Supplier, email và quy tắc xử lý: [`Config/suppliers.csv`](Config/suppliers.csv)
- Danh sách tool trong launcher: [`launcher-config.json`](launcher-config.json)

Đường dẫn tương đối được tính từ thư mục cài đặt. Có thể đặt biến môi trường `AUTOTOOLS_ROOT` nếu cần chạy script từ vị trí khác.

> `Config/suppliers.csv` chứa thông tin liên hệ nghiệp vụ. Repo phải được giữ **private** và không được đính kèm file này vào issue, log hoặc ảnh chụp chưa che dữ liệu.

## Kiểm tra

Chạy toàn bộ kiểm tra không cần Excel hoặc Outlook:

```powershell
Get-ChildItem ".\Tests\Test-*.ps1" | Sort-Object Name | ForEach-Object {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $_.FullName
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $($_.Name)" }
}
```

GitHub Actions cũng tự động parse toàn bộ script và chạy test trên mỗi push/PR vào `main`.

## Build bộ cài

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\Build-Installer.ps1"
```

Artifact được tạo tại `dist\AutoTools-Setup-1.2.0.exe`. Version trong launcher, build script và Inno Setup phải luôn khớp nhau; `Tests/Test-Installer.ps1` kiểm tra hợp đồng này.

## Cấu trúc repo

```text
Auto-Tools/
├── Config/                 # Đường dẫn và supplier master
├── Input/                  # Chỉ chứa file input mẫu được phép commit
├── Installer/              # Inno Setup manifest
├── Scripts/                # Các workflow PowerShell
├── Tests/                  # Test chạy độc lập bằng PowerShell 5.1
├── Launcher.ps1            # Giao diện chính
└── launcher-config.json    # Danh sách tool trong launcher
```

Dữ liệu vận hành trong `Input`, `Output`, `Logs`, `Temp`, `dist` và các thư mục output con không được commit.

## Tài liệu và hỗ trợ

- [Hướng dẫn sử dụng đầy đủ](HUONG_DAN_SETUP_AUTO_TOOLS.txt)
- [Lịch sử thay đổi](CHANGELOG.md)
- [Hướng dẫn đóng góp](CONTRIBUTING.md)
- [Chính sách bảo mật](SECURITY.md)

Khi báo lỗi, hãy cung cấp tên tool, bước tái hiện và đoạn log đã che dữ liệu nhạy cảm. Không đính kèm workbook, danh sách supplier hoặc email thật.

## Phạm vi sử dụng

Repo chưa công bố giấy phép mã nguồn mở. Hãy liên hệ maintainer trước khi sao chép hoặc phân phối ngoài phạm vi được cấp quyền.
