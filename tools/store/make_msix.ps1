<#
  Builds the Microsoft Store package (MSIX) of Microads Office from the published single executable.

  The Store re-signs the package with Microsoft's certificate after certification: users installing from the
  Store never see SmartScreen or browser download warnings. Identity values come from Partner Center
  (Apps and games › Microads Office › Product management › Product identity).

  Usage:
    powershell -ExecutionPolicy Bypass -File tools\store\make_msix.ps1 `
        -IdentityName "12345Microads.MicroadsOffice" -Publisher "CN=XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX" `
        -PublisherDisplay "Công ty TNHH Công nghệ Microads"
    (without parameters a local test identity is used; the .msix is then only for testing)
#>
param(
  [string]$IdentityName = 'MicroadsOffice.LocalTest',
  [string]$Publisher = 'CN=Microads Office Local Test',
  [string]$PublisherDisplay = 'Công ty TNHH Công nghệ Microads',
  [string]$Exe,
  [switch]$RegisterForTest
)
$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path $PSScriptRoot '..\..')
if (-not $Exe) { $Exe = Join-Path $root 'release\MicroadsOffice\MicroadsOffice.exe' }
if (-not (Test-Path $Exe)) { throw "Published executable not found: $Exe (run build.ps1 first)" }

$sdk = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin' -Directory -ErrorAction SilentlyContinue |
  Where-Object { Test-Path (Join-Path $_.FullName 'x64\makeappx.exe') } | Sort-Object Name -Descending | Select-Object -First 1
if (-not $sdk) { throw 'Windows SDK (makeappx.exe) not found' }
$makeappx = Join-Path $sdk.FullName 'x64\makeappx.exe'
$makepri = Join-Path $sdk.FullName 'x64\makepri.exe'

$v = (Get-Item $Exe).VersionInfo
$version = "{0}.{1}.{2}.0" -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart

$work = Join-Path $root '.build\msix'
$layout = Join-Path $work 'layout'
if (Test-Path $layout) { Remove-Item $layout -Recurse -Force }
New-Item -ItemType Directory -Force -Path "$layout\Assets" | Out-Null
Copy-Item $Exe "$layout\MicroadsOffice.exe"

# logos rendered from the vector brand art inside the executable itself
$p = Start-Process -FilePath "$layout\MicroadsOffice.exe" -ArgumentList '--export-store-assets', "`"$layout\Assets`"" -Wait -PassThru -NoNewWindow
if ($p.ExitCode -ne 0) { throw 'asset export failed' }

$manifest = Get-Content (Join-Path $PSScriptRoot 'AppxManifest.template.xml') -Raw -Encoding UTF8
$manifest = $manifest.Replace('{{IDENTITY_NAME}}', $IdentityName).Replace('{{PUBLISHER}}', [Security.SecurityElement]::Escape($Publisher)).
  Replace('{{PUBLISHER_DISPLAY}}', [Security.SecurityElement]::Escape($PublisherDisplay)).Replace('{{VERSION}}', $version)
[IO.File]::WriteAllText("$layout\AppxManifest.xml", $manifest, (New-Object Text.UTF8Encoding($false)))

# resources.pri lets Windows pick the right icon size (taskbar, Start, file icons)
Push-Location $work
try {
  & $makepri createconfig /cf priconfig.xml /dq vi-VN_en-US /pv 10.0.0 /o | Out-Null
  & $makepri new /pr $layout /cf priconfig.xml /of "$layout\resources.pri" /o | Out-Null
  if ($LASTEXITCODE -ne 0) { throw 'makepri failed' }
} finally { Pop-Location }

$out = Join-Path $root "release\store"
New-Item -ItemType Directory -Force -Path $out | Out-Null
$suffix = if ($IdentityName -eq 'MicroadsOffice.LocalTest') { '_LOCALTEST-do-not-upload' } else { '' }
$msix = Join-Path $out "MicroadsOffice_$($version)_x64$suffix.msix"
& $makeappx pack /d $layout /p $msix /o | Select-Object -Last 2
if ($LASTEXITCODE -ne 0) { throw 'makeappx failed' }
Write-Host "MSIX: $msix ($([math]::Round((Get-Item $msix).Length / 1MB, 1)) MB) — identity $IdentityName, version $version"

if ($RegisterForTest) {
  # Developer Mode: register the unpacked layout (files stay on E:) to test the Store edition locally.
  Add-AppxPackage -Register "$layout\AppxManifest.xml" -ForceApplicationShutdown
  Get-AppxPackage -Name $IdentityName | Select-Object Name, PackageFamilyName, Version
}
