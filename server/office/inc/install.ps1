# Microads Office — cài đặt bằng PowerShell (không qua trình duyệt)
#   irm https://microads.vn/office/install.ps1 | iex
# Tải bản mới nhất từ microads.vn, kiểm tra SHA-256 và chữ ký số ECDSA của Microads rồi mới mở trình cài đặt.
# Tùy chọn (đặt trước khi chạy lệnh):
#   $env:MICROADS_SILENT = '1'                        cài im lặng, không hỏi (cài cho người dùng hiện tại)
#   $env:MICROADS_INSTALL_DIR = 'D:\Apps\Microads Office'   thư mục cài đặt
# Microads Office — install with PowerShell (no browser). Downloads the latest release from microads.vn,
# verifies its SHA-256 and the Microads ECDSA signature, then starts the setup. Windows PowerShell 5.1 and PowerShell 7.
# Run it with "irm ... | iex" (served as UTF-8 without BOM: a saved copy needs UTF-8 with BOM for Windows PowerShell 5.1).
& {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    $manifestUrl = '__MANIFEST_URL__'
    $o = $env:MICROADS_UPDATE_URL; $u = $null
    if ($o -and [Uri]::TryCreate($o, [UriKind]::Absolute, [ref]$u) -and ($u.Scheme -eq 'https' -or $u.IsLoopback)) { $manifestUrl = $o }
    # Public half of the Microads release-signing key (ECDSA P-256) — the same key the apps use for their updates.
    $publicKey = 'MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEXSjzXUOQ/LgAupMzJCQNytd+BZaRwafJixpx9FX1LW4jyHov1tOAxAi86oetJjZFXeVbutcMCa52lXxcX42CXw=='
    $vi = ((Get-Culture).TwoLetterISOLanguageName -eq 'vi') -or ((Get-UICulture).TwoLetterISOLanguageName -eq 'vi') -or ([TimeZoneInfo]::Local.Id -eq 'SE Asia Standard Time')
    if ($env:MICROADS_LANG -in @('vi', 'en')) { $vi = $env:MICROADS_LANG -eq 'vi' }
    function Say([string]$textVi, [string]$textEn) { if ($vi) { $textVi } else { $textEn } }

    function ConvertFrom-Der([byte[]]$d) {
        # DER SEQUENCE { INTEGER r, INTEGER s } -> r||s (64 bytes)
        if ($d.Length -lt 8 -or $d[0] -ne 0x30) { return $null }
        $p = 2; if ($d[1] -eq 0x81) { $p = 3 }
        $out = New-Object byte[] 64
        for ($k = 0; $k -lt 2; $k++) {
            if ($p + 2 -gt $d.Length -or $d[$p] -ne 0x02) { return $null }
            $l = [int]$d[$p + 1]; $p += 2
            if ($l -lt 1 -or $l -gt 33 -or $p + $l -gt $d.Length) { return $null }
            $start = $p; $cnt = $l
            while ($cnt -gt 32 -and $d[$start] -eq 0) { $start++; $cnt-- }
            if ($cnt -gt 32) { return $null }
            [Array]::Copy($d, $start, $out, $k * 32 + (32 - $cnt), $cnt)
            $p += $l
        }
        if ($p -ne $d.Length) { return $null }
        return , $out
    }

    function Test-ReleaseSignature([string]$version, [string]$sha, [string]$signature) {
        try {
            $spki = [Convert]::FromBase64String($publicKey)
            $sig = ConvertFrom-Der ([Convert]::FromBase64String($signature))
            if ($null -eq $sig) { return $false }
            $msg = [Text.Encoding]::UTF8.GetBytes("MicroadsOffice|$version|$sha")
            if ($PSVersionTable.PSVersion.Major -ge 6) {
                $ec = [Security.Cryptography.ECDsa]::Create(); $read = 0
                $ec.ImportSubjectPublicKeyInfo($spki, [ref]$read)
                return $ec.VerifyData($msg, $sig, [Security.Cryptography.HashAlgorithmName]::SHA256)
            }
            $prefix = [Convert]::ToBase64String($spki[0..26])
            if ($spki.Length -ne 91 -or $prefix -ne 'MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE') { return $false }
            [byte[]]$blob = @(0x45, 0x43, 0x53, 0x31, 0x20, 0, 0, 0) + $spki[27..90]   # BCRYPT_ECDSA_PUBLIC_P256 "ECS1"
            $key = [Security.Cryptography.CngKey]::Import($blob, [Security.Cryptography.CngKeyBlobFormat]::EccPublicBlob)
            $ec = New-Object Security.Cryptography.ECDsaCng($key)
            $hash = (New-Object Security.Cryptography.SHA256Cng).ComputeHash($msg)
            return $ec.VerifyHash($hash, $sig)
        } catch { return $false }
    }

    if (-not [Environment]::Is64BitOperatingSystem -or [Environment]::OSVersion.Version.Major -lt 10) {
        Write-Host (Say 'Microads Office cần Windows 10 hoặc Windows 11 bản 64-bit.' 'Microads Office needs 64-bit Windows 10 or Windows 11.') -ForegroundColor Red
        return
    }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    try { if ([Environment]::OSVersion.Version.Build -ge 22000) { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]'Tls13' } } catch { }

    # Install folder (optional): full path only — a relative one would be resolved against the temporary download folder.
    $installDir = $null
    if ($env:MICROADS_INSTALL_DIR) {
        $d = [Environment]::ExpandEnvironmentVariables($env:MICROADS_INSTALL_DIR.Trim().Trim('"'))
        if ($d -match '"|%' -or -not ($d -match '^[A-Za-z]:[\\/]' -or $d -match '^\\\\[^?.]')) {
            Write-Host ((Say 'MICROADS_INSTALL_DIR phải là đường dẫn đầy đủ, ví dụ D:\Apps\Microads Office (đang là "{0}").' 'MICROADS_INSTALL_DIR must be a full path such as D:\Apps\Microads Office (got "{0}").') -f $env:MICROADS_INSTALL_DIR) -ForegroundColor Red
            return
        }
        $installDir = [IO.Path]::GetFullPath($d)
        if ($installDir -ne [IO.Path]::GetPathRoot($installDir)) { $installDir = $installDir.TrimEnd('\', '/') }
    }

    $work = $null; $lock = $null
    try {
        Write-Host (Say 'Microads Office — đang tìm phiên bản mới nhất tại microads.vn...' 'Microads Office — looking for the latest version on microads.vn...') -ForegroundColor Cyan
        $m = Invoke-RestMethod -UseBasicParsing -Uri ($manifestUrl + '?v=ps-install') -Headers @{ 'Cache-Control' = 'no-cache' }
        $r = $m.latest
        if ($null -eq $r) { throw (Say 'Chưa có phiên bản nào được phát hành.' 'No version has been released yet.') }
        $ver = [string]$r.version; $url = [string]$r.url; $sha = ([string]$r.sha256).ToLowerInvariant(); $size = 0L
        $dl = $null
        if ($ver -cnotmatch '^[0-9]{1,4}\.[0-9]{1,4}\.[0-9]{1,6}\z' -or $sha -cnotmatch '^[0-9a-f]{64}\z' -or -not [long]::TryParse([string]$r.size, [ref]$size) -or
            $size -le 0 -or $size -gt 1GB -or -not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$dl) -or
            -not ($dl.Scheme -eq 'https' -or ($dl.Scheme -eq 'http' -and $dl.IsLoopback))) {
            throw (Say 'Thông tin phiên bản không hợp lệ — đã dừng để bảo vệ máy tính.' 'Invalid release information — stopped to protect this computer.')
        }
        $work = Join-Path ([IO.Path]::GetTempPath()) ('MicroadsOffice-Setup-' + [Guid]::NewGuid().ToString('N').Substring(0, 10))
        New-Item -ItemType Directory -Path $work | Out-Null
        $file = Join-Path $work "MicroadsOffice_Setup_$ver.exe"
        Write-Host ((Say 'Đang tải Microads Office {0} ({1:N1} MB)...' 'Downloading Microads Office {0} ({1:N1} MB)...') -f $ver, ($size / 1MB))
        Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $file
        if ((Get-Item -LiteralPath $file).Length -ne $size) { throw (Say 'Tệp tải về không đủ dung lượng — vui lòng thử lại.' 'The download is incomplete — please try again.') }

        # Keep the file open read-only (nobody can change it) from the check until Windows has started it.
        $lock = [IO.File]::Open($file, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $actual = (Get-FileHash -InputStream $lock -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actual -ne $sha -or -not (Test-ReleaseSignature $ver $actual ([string]$r.signature))) {
            throw (Say 'Tệp tải về không có chữ ký số hợp lệ của Microads — đã hủy để bảo vệ máy tính.' 'The download does not carry a valid Microads signature — cancelled to protect this computer.')
        }
        Write-Host (Say 'Chữ ký số của Microads hợp lệ ✓' 'Valid Microads digital signature ✓') -ForegroundColor Green

        $silent = $env:MICROADS_SILENT -eq '1'
        $setupArgs = '--setup --lang=' + $(if ($vi) { 'vi' } else { 'en' })
        if ($silent) { $setupArgs += ' --silent' }
        if ($installDir) { $setupArgs += ' --dir "' + $installDir + $(if ($installDir.EndsWith('\')) { '\' } else { '' }) + '"' }
        if (-not $silent) { Write-Host (Say 'Đang mở trình cài đặt — làm theo hướng dẫn trên màn hình.' 'Starting the setup — follow the instructions on screen.') }
        else { Write-Host (Say 'Đang cài đặt...' 'Installing...') }
        $proc = Start-Process -FilePath $file -ArgumentList $setupArgs -PassThru
        $lock.Dispose(); $lock = $null
        $proc.WaitForExit()
        if ($proc.ExitCode -eq 0) {
            Write-Host ((Say 'Đã cài Microads Office {0}. Mở Word, Excel hoặc PDF từ menu Start.' 'Microads Office {0} is installed. Open Word, Excel or PDF from the Start menu.') -f $ver) -ForegroundColor Green
        } elseif ($silent) {
            Write-Host ((Say 'Cài đặt không thành công (mã {0}). Chi tiết: %TEMP%\MicroadsOffice_Setup.log' 'Installation failed (code {0}). Details: %TEMP%\MicroadsOffice_Setup.log') -f $proc.ExitCode) -ForegroundColor Red
        } else {
            Write-Host (Say 'Trình cài đặt đã đóng.' 'The setup was closed.')
        }
    }
    catch {
        Write-Host ((Say 'Không cài được Microads Office: {0}' 'Could not install Microads Office: {0}') -f $_.Exception.Message) -ForegroundColor Red
        Write-Host (Say 'Bạn có thể tải bộ cài tại https://microads.vn/office' 'You can download the installer from https://microads.vn/office')
    }
    finally {
        if ($lock) { $lock.Dispose() }
        if ($work) { for ($i = 0; $i -lt 10 -and (Test-Path $work); $i++) { try { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 500 } } }
    }
}
