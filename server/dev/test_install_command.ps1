# Integration test: PowerShell install command + web installer against the local test site (server/dev/test_site.py --keep-running).
#   powershell -NoProfile -ExecutionPolicy Bypass -File server\dev\test_install_command.ps1 -Work <folder> -WebSetup <MicroadsOffice_Setup.exe> [-Shell <powershell.exe|pwsh.exe>]
# Everything is installed inside <Work> in sandbox mode (shortcuts / registry redirected, removed at the end).
param([Parameter(Mandatory)] [string]$Work, [Parameter(Mandatory)] [string]$WebSetup, [string]$Shell = 'powershell.exe', [string]$Base = 'http://127.0.0.1:8099/office')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$pass = 0; $fail = 0
function Check([bool]$ok, [string]$what) { if ($ok) { $script:pass++; Write-Host "PASS  $what" } else { $script:fail++; Write-Host "FAIL  $what" -ForegroundColor Red } }
if (Test-Path $Work) { Remove-Item $Work -Recurse -Force }
New-Item -ItemType Directory -Path "$Work\tmp" | Out-Null
$data = Join-Path (Split-Path $Work) 'server-dev\data'
$sandbox = "$Work\sandbox"
$common = @{ TEMP = "$Work\tmp"; TMP = "$Work\tmp"; MICROADS_SETUP_SANDBOX = $sandbox; MICROADS_DATA_DIR = "$Work\appdata"; MICROADS_LANG = 'vi' }

function Run-Command([hashtable]$extra, [switch]$NoWait) {
    $psi = New-Object Diagnostics.ProcessStartInfo $Shell
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"[Console]::OutputEncoding=[Text.Encoding]::UTF8; irm $Base/install.ps1 | iex`""
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    foreach ($k in $common.Keys) { $psi.EnvironmentVariables[$k] = $common[$k] }
    foreach ($k in $extra.Keys) { $psi.EnvironmentVariables[$k] = $extra[$k] }
    $p = [Diagnostics.Process]::Start($psi)
    if ($NoWait) { return $p }
    $out = $p.StandardOutput.ReadToEnd() + $p.StandardError.ReadToEnd(); $p.WaitForExit()
    return $out
}
$shellName = [IO.Path]::GetFileNameWithoutExtension($Shell)

# ---------------------------------------------------------------- 1. silent install
$inst = "$Work\inst-ps"
$out = Run-Command @{ MICROADS_SILENT = '1'; MICROADS_INSTALL_DIR = '%MO_PS_TEST%\inst-ps'; MO_PS_TEST = $Work }
Write-Host ($out.Trim() -replace '(?m)^', '      ')
Check ($out -match 'Đang tải Microads Office 2\.0\.1') "[$shellName] Vietnamese text decoded correctly (UTF-8)"
Check ($out -match 'Chữ ký số của Microads hợp lệ') "[$shellName] Microads signature verified"
Check ($out -match 'Đã cài Microads Office 2\.0\.1') "[$shellName] silent install reports success"
Check (Test-Path "$inst\MicroadsOffice.exe") "[$shellName] program installed into MICROADS_INSTALL_DIR (%VARIABLE% expanded)"
Check (-not (Get-Item "$inst\MicroadsOffice.exe" -Stream Zone.Identifier -ErrorAction SilentlyContinue)) "[$shellName] installed program has no Mark of the Web"
Check (@(Get-ChildItem "$Work\tmp" -Directory -Filter 'MicroadsOffice-Setup-*').Count -eq 0) "[$shellName] temporary download removed"
$stats = Get-Content "$data\stats\$(Get-Date -Format yyyy-MM-dd).json" -Raw | ConvertFrom-Json
Check ($stats.'ps-install' -ge 1) "[$shellName] PowerShell installs counted on the server (ps-install)"

# ---------------------------------------------------------------- 1b. relative install folder refused
$out = Run-Command @{ MICROADS_SILENT = '1'; MICROADS_INSTALL_DIR = 'Microads Office' }
Check ($out -match 'phải là đường dẫn đầy đủ') "[$shellName] relative MICROADS_INSTALL_DIR refused with an explanation"
Check (@(Get-ChildItem "$Work\tmp" -Directory -Filter 'MicroadsOffice-Setup-*').Count -eq 0) "[$shellName] nothing downloaded for a refused folder"

# ---------------------------------------------------------------- 2. tampered release information
$cfgPath = "$data\config.json"; $cfgRaw = [IO.File]::ReadAllText($cfgPath)
try {
    $cfg = $cfgRaw | ConvertFrom-Json
    $cfg.releases[0].sha256 = ('0' * 64)
    [IO.File]::WriteAllText($cfgPath, ($cfg | ConvertTo-Json -Depth 10))
    $out = Run-Command @{ MICROADS_SILENT = '1'; MICROADS_INSTALL_DIR = "$Work\inst-bad" }
    Check ($out -match 'không có chữ ký số hợp lệ của Microads') "[$shellName] file not matching the published SHA-256 is refused"
    Check (-not (Test-Path "$Work\inst-bad\MicroadsOffice.exe")) "[$shellName] nothing installed"
    Check (@(Get-ChildItem "$Work\tmp" -Directory -Filter 'MicroadsOffice-Setup-*').Count -eq 0) "[$shellName] refused download deleted"
} finally { [IO.File]::WriteAllText($cfgPath, $cfgRaw) }

# ---------------------------------------------------------------- 3. interactive: the setup window opens, Cancel closes it
$p = Run-Command @{} -NoWait
$root = [Windows.Automation.AutomationElement]::RootElement
$win = $null
for ($i = 0; $i -lt 120 -and -not $win; $i++) {
    Start-Sleep -Milliseconds 500
    foreach ($proc in Get-Process -Name 'MicroadsOffice_Setup_2.0.1' -ErrorAction SilentlyContinue) {
        try { if ($proc.Path -like "$Work\*" -and $proc.MainWindowHandle -ne 0) { $win = [Windows.Automation.AutomationElement]::FromHandle($proc.MainWindowHandle) } } catch { }
    }
}
Check ($null -ne $win) "[$shellName] interactive: the Microads Office setup window opens (no browser, no SmartScreen)"
if ($win) {
    $cancel = $win.FindFirst([Windows.Automation.TreeScope]::Descendants, (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, 'Hủy')))
    Check ($null -ne $cancel) "[$shellName] interactive: setup shown in Vietnamese (Hủy button)"
    if ($cancel) { $cancel.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() }
}
$out = $p.StandardOutput.ReadToEnd(); $p.WaitForExit(60000) | Out-Null
Check ($out -match 'Trình cài đặt đã đóng') "[$shellName] interactive: closing the setup is reported"
Check (@(Get-ChildItem "$Work\tmp" -Directory -Filter 'MicroadsOffice-Setup-*').Count -eq 0) "[$shellName] interactive: temporary download removed"

# ---------------------------------------------------------------- 4. web installer against the real PHP manifest
$inst2 = "$Work\inst-web"
$psi = New-Object Diagnostics.ProcessStartInfo $WebSetup, "--silent --dir `"$inst2`""
$psi.UseShellExecute = $false
foreach ($k in $common.Keys) { $psi.EnvironmentVariables[$k] = $common[$k] }
$psi.EnvironmentVariables['MICROADS_UPDATE_URL'] = "$Base/api/manifest.json"
$w = [Diagnostics.Process]::Start($psi); $w.WaitForExit(180000) | Out-Null
Check ($w.ExitCode -eq 0 -and (Test-Path "$inst2\MicroadsOffice.exe")) "web installer + PHP manifest: silent install succeeds (code $($w.ExitCode))"

# ---------------------------------------------------------------- cleanup
foreach ($d in $inst, $inst2) {
    if (Test-Path "$d\MicroadsOffice.exe") {
        $psi = New-Object Diagnostics.ProcessStartInfo "$d\MicroadsOffice.exe", '--uninstall --silent'; $psi.UseShellExecute = $false
        foreach ($k in $common.Keys) { $psi.EnvironmentVariables[$k] = $common[$k] }
        [Diagnostics.Process]::Start($psi).WaitForExit(30000) | Out-Null
    }
}
Start-Sleep -Seconds 4
Check (-not (Test-Path "$inst\MicroadsOffice.exe") -and -not (Test-Path "$inst2\MicroadsOffice.exe")) "test installations uninstalled"
Remove-Item 'HKCU:\Software\MicroadsOfficeSetupSandbox' -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "`nTOTAL: $pass passed, $fail failed"
exit $(if ($fail) { 1 } else { 0 })
