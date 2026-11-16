# Microads Office trên winget (Windows Package Manager) — miễn phí, không phải Microsoft Store

Sau khi được duyệt, khách và bộ phận IT cài bằng một lệnh, **không có cảnh báo của trình duyệt hay SmartScreen**:

```
winget install --id Microads.Office -e
```

winget có sẵn trên Windows 10 (1809+) và Windows 11. Nó tải bộ cài đầy đủ từ microads.vn, kiểm tra SHA-256 ghi trong hồ sơ, rồi cài im lặng cho người dùng hiện tại (`--install --silent`, không cần quyền quản trị).

## Hồ sơ đã chuẩn bị

`manifests/m/Microads/Office/<phiên bản>/`: 4 tệp YAML (version, installer, locale en-US, locale vi-VN), schema 1.10.0.
Tạo cho phiên bản mới: `build.ps1 -Publish` tự chạy, hoặc chạy tay
`python -I tools\winget\make_manifests.py release\MicroadsOffice_Setup_x.y.z.exe.sig.json`.
Kiểm tra: `winget validate --manifest tools\winget\manifests\m\Microads\Office\x.y.z` (2.0.1: *Manifest validation succeeded*).

## Gửi lần đầu (khoảng 10 phút, cần tài khoản GitHub của Microads)

1. Vào https://github.com/microsoft/winget-pkgs → **Fork**.
2. Trong bản fork, tạo nhánh mới và thêm thư mục `manifests/m/Microads/Office/2.0.1/` với 4 tệp trong thư mục này (GitHub web: *Add file › Upload files*, giữ nguyên đường dẫn).
3. Mở **Pull request** vào `microsoft/winget-pkgs` (`master`), tiêu đề `New package: Microads.Office version 2.0.1`, đánh dấu các ô trong mẫu PR.
4. Hệ thống của Microsoft tự kiểm tra: tải bộ cài, quét Defender, cài thử im lặng trong máy ảo. Sau đó người duyệt xem xét. Thường mất 1–7 ngày; nếu có nhãn lỗi, làm theo hướng dẫn trong PR.
5. Khi PR được gộp: vào **microads.vn/office/admin › Nội dung quảng bá › Windows Package Manager (winget)** → nhập `Microads.Office` → *Lưu*. Trang tải sẽ hiện lệnh `winget install` kèm nút Sao chép.

Cách nhanh hơn (tùy chọn): công cụ chính thức `wingetcreate` (`winget install wingetcreate`), rồi `wingetcreate submit --token <GitHub PAT> tools\winget\manifests\m\Microads\Office\2.0.1`.

## Mỗi phiên bản mới

`build.ps1 -Publish` tạo thư mục hồ sơ của phiên bản mới. Gửi một PR `Update Microads.Office to x.y.z` (hoặc `wingetcreate update Microads.Office -v x.y.z -u https://microads.vn/office/download/MicroadsOffice_Setup_x.y.z.exe --submit`).
Bộ cài các bản cũ phải **giữ nguyên trên máy chủ** vì hồ sơ đã duyệt trỏ tới chúng (trang quản trị chỉ xóa khi thật cần).

## Hợp đồng kỹ thuật (đừng phá)

- Bộ cài đầy đủ phải tiếp tục nhận `--install --silent --dir "<thư mục>"` và trả mã thoát 0 khi thành công.
- Mục gỡ cài đặt của Windows giữ khóa `HKCU\...\Uninstall\MicroadsOffice` (= `ProductCode` trong hồ sơ) và `DisplayVersion` = phiên bản, để `winget upgrade` / `winget uninstall` nhận ra.
