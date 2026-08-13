# Changelog

Các thay đổi đáng chú ý của Auto Tools được ghi tại đây theo định dạng [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) và Semantic Versioning.

## [Unreleased]

## [1.2.0] - 2026-08-13

### Added

- Launcher responsive theo DPI, có trạng thái `READY`/`MISSING` và accessible name cho các thao tác chính.
- Supplier Commitment có fallback quét cây thư mục Outlook khi đường dẫn BU thay đổi.
- Nhận diện BU/supplier từ tiêu đề ISPIH, đối chiếu với `Config/suppliers.csv` và tạo Outlook reminder draft cho supplier còn thiếu.
- Process PO cho phép chọn chia theo supplier hoặc gộp theo site TC5/TN5.
- AutoImport archive thư mục nguồn vào `Older` sau khi tạo dữ liệu ngày mới.
- Test tự động cho launcher, installer, Supplier Commitment, Process PO và AutoImport.
- GitHub Actions chạy syntax check và test trên mỗi push/PR vào `main`.

### Changed

- Đồng bộ version launcher, build script và installer lên `1.2.0`.
- Loại bỏ đường dẫn user hard-code khỏi AutoImport để chạy được trên máy khác.
- Chuẩn hóa README, tài liệu sử dụng và thông tin release.

### Security

- Bổ sung hướng dẫn không commit hoặc chia sẻ dữ liệu vận hành và thông tin liên hệ supplier.

## [1.1.0] - 2026-07-20

### Added

- Baseline đầu tiên của bộ Auto Tools hoàn chỉnh trên GitHub.

[Unreleased]: https://github.com/hugo-lcqh/Auto-Tools/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/hugo-lcqh/Auto-Tools/releases/tag/v1.2.0
[1.1.0]: https://github.com/hugo-lcqh/Auto-Tools/commit/8d89997
