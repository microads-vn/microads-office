<#
  Microads Office — build / test / package / sign / publish.
  Everything (SDK, NuGet cache, temp, tool profiles) stays inside E:\Office\.build. Nothing is written to C: or D:.

  Usage:
    powershell -ExecutionPolicy Bypass -File build.ps1                    # build + tests + single-exe package + signed setup
    powershell -ExecutionPolicy Bypass -File build.ps1 -UiTests           # also run the PDF end-to-end UI tests (uses mouse/keyboard)
    powershell -ExecutionPolicy Bypass -File build.ps1 -Publish [-ForceUpdate] [-NotesVi notes_vi.txt] [-NotesEn notes_en.txt]
          # uploads the signed setup to microads.vn/office; -ForceUpdate = "Cập nhật toàn bộ" (every installation updates itself)
  The version comes from <Version> in src\MicroadsOffice\MicroadsOffice.csproj.

  Authenticode (removes "unknown publisher"; SmartScreen/Chrome warnings fade as reputation builds):
    build.ps1 -CodeSignThumbprint <SHA1 thumbprint of the OV code-signing certificate on its USB token / cloud HSM>
    (or set MICROADS_CODESIGN_THUMBPRINT). The token PIN is asked by the CA driver.
  Microsoft Store package (no SmartScreen / browser warnings at all, signed by Microsoft):
    build.ps1 -Store -StoreIdentityName MicroadsTech.MicroadsOffice -StorePublisher "CN=1BE10500-500A-4011-8C02-028DAF5EF854" [-StorePublisherDisplay "Microads Tech"]
    (values: Partner Center › Apps and games › Microads Office › Product identity; the display name must match it exactly)

  Web installer (release\web-setup\MicroadsOffice_Setup.exe, the main download button on microads.vn/office):
    a small program that NEVER changes and always installs the latest signed release, so the reputation it earns with
    Chrome / SmartScreen (tied to its SHA-256) keeps growing across releases. Every build rebuilds it and compares it with
    the frozen copy; the frozen copy is only replaced with -NewWebSetup (= reputation starts again from zero).
    build.ps1 -WebSetupOnly [-UiTests] [-Publish]   # only the web installer (+ its end-to-end test / upload)
  winget: -Publish also writes tools\winget\manifests\...\<version> for the Windows Package Manager pull request.
#>
param([switch]$UiTests, [switch]$Publish, [switch]$ForceUpdate, [string]$NotesVi, [string]$NotesEn, [switch]$SkipTests,
      [string]$CodeSignThumbprint = $env:MICROADS_CODESIGN_THUMBPRINT, [switch]$Store, [string]$StoreIdentityName, [string]$StorePublisher, [string]$StorePublisherDisplay = 'Microads Tech',
      [switch]$WebSetupOnly, [switch]$NewWebSetup)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$b = Join-Path $root '.build'
New-Item -ItemType Directory -Force -Path "$b\tmp", "$b\home\AppData\Roaming", "$b\home\AppData\Local", "$b\nuget\packages" | Out-Null

$env:DOTNET_ROOT = "$b\dotnet"
$env:PATH = "$b\dotnet;$env:PATH"
$env:DOTNET_CLI_HOME = "$b\home"; $env:HOME = "$b\home"
$origProfile = $env:USERPROFILE; $origAppData = $env:APPDATA; $origLocal = $env:LOCALAPPDATA
$env:USERPROFILE = "$b\home"; $env:APPDATA = "$b\home\AppData\Roaming"; $env:LOCALAPPDATA = "$b\home\AppData\Local"
$env:NUGET_PACKAGES = "$b\nuget\packages"; $env:NUGET_HTTP_CACHE_PATH = "$b\nuget\http-cache"; $env:NUGET_PLUGINS_CACHE_PATH = "$b\nuget\plugins-cache"
$env:TEMP = "$b\tmp"; $env:TMP = "$b\tmp"; $env:DOTNET_BUNDLE_EXTRACT_BASE_DIR = "$b\tmp"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'; $env:DOTNET_NOLOGO = '1'; $env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:DOTNET_GENERATE_ASPNET_CERTIFICATE = 'false'; $env:DOTNET_ADD_GLOBAL_TOOLS_TO_PATH = 'false'
$env:MSBUILDDISABLENODEREUSE = '1'; $env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'; $env:UseSharedCompilation = 'false'
$env:PYTHONDONTWRITEBYTECODE = '1'; $env:PYTHONIOENCODING = 'utf-8'
$py = "$b\pyenv\Scripts\python.exe"

if (-not (Test-Path "$b\dotnet\dotnet.exe")) {
  Write-Host 'Installing portable .NET 10 SDK into .build\dotnet ...'
  Invoke-WebRequest -UseBasicParsing -Uri 'https://dot.net/v1/dotnet-install.ps1' -OutFile "$b\dotnet-install.ps1"
  & "$b\dotnet-install.ps1" -Channel 10.0 -InstallDir "$b\dotnet" -NoPath
}

[xml]$proj = Get-Content "$root\src\MicroadsOffice\MicroadsOffice.csproj"
$version = $proj.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
Write-Host "Microads Office $version" -ForegroundColor Cyan

$setup = Join-Path $root "release\MicroadsOffice_Setup_$version.exe"
$webFrozen = Join-Path $root 'release\web-setup\MicroadsOffice_Setup.exe'

Push-Location $root
try {
  if (-not $WebSetupOnly) {
  dotnet build MicroadsOffice.sln -c Release
  if ($LASTEXITCODE -ne 0) { throw 'Build failed' }

  if (-not $SkipTests) {
    Write-Host "`n== PDF core tests ==" -ForegroundColor Cyan
    & "$root\tests\MicroadsOffice.Tests\bin\Release\net10.0-windows\MicroadsOffice.Tests.exe" core "$b\test-out"
    if ($LASTEXITCODE -ne 0) { throw 'PDF core tests failed' }
    Write-Host "`n== Word tests ==" -ForegroundColor Cyan
    $p = Start-Process -FilePath "$root\tests\MicroadsOffice.Word.Tests\bin\Release\net10.0-windows\MicroadsOffice.Word.Tests.exe" -ArgumentList "all `"$b\word-test`" vi" -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) { throw 'Word tests failed' }
    Write-Host "`n== Excel tests ==" -ForegroundColor Cyan
    $p = Start-Process -FilePath "$root\tests\MicroadsOffice.Excel.Tests\bin\Release\net10.0-windows\MicroadsOffice.Excel.Tests.exe" -ArgumentList "all `"$b\excel-test`"" -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) { throw 'Excel tests failed' }
    Write-Host "`n== AI assistant tests (local fake server: sign-in, providers, MCP, Word/Excel panes) ==" -ForegroundColor Cyan
    foreach ($m in 'core', 'word', 'excel', 'pdf') {
      $p = Start-Process -FilePath "$root\tests\MicroadsOffice.AI.Tests\bin\Release\net10.0-windows\MicroadsOffice.AI.Tests.exe" -ArgumentList "$m `"$b\ai-test`" vi" -Wait -PassThru -NoNewWindow
      if ($p.ExitCode -ne 0) { throw "AI tests ($m) failed" }
    }
  }

  Write-Host "`n== Package: one self-contained executable ==" -ForegroundColor Cyan
  $out = Join-Path $root 'release\MicroadsOffice'
  if (Test-Path $out) { Get-ChildItem $out -Exclude 'Data' | Remove-Item -Recurse -Force }
  dotnet publish src\MicroadsOffice -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=false -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none -p:DebugSymbols=false -o $out
  if ($LASTEXITCODE -ne 0) { throw 'Publish failed' }
  # PDFium is embedded in the executable and extracted on first use: no separate DLL is shipped.
  Get-ChildItem $out -File | Where-Object { $_.Name -ne 'MicroadsOffice.exe' } | Remove-Item -Force

  if ($CodeSignThumbprint) {
    Write-Host "`n== Authenticode signing ==" -ForegroundColor Cyan
    $sdkBin = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin' -Directory -ErrorAction SilentlyContinue |
      Where-Object { Test-Path (Join-Path $_.FullName 'x64\signtool.exe') } | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $sdkBin) { throw 'signtool.exe (Windows SDK) not found' }
    $signtool = Join-Path $sdkBin.FullName 'x64\signtool.exe'
    # SHA-256 file digest + RFC 3161 timestamp: the signature stays valid after the certificate expires.
    & $signtool sign /sha1 $CodeSignThumbprint /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 `
      /d "Microads Office" /du "https://microads.vn" "$out\MicroadsOffice.exe"
    if ($LASTEXITCODE -ne 0) { throw 'Authenticode signing failed' }
    & $signtool verify /pa /q "$out\MicroadsOffice.exe"
    if ($LASTEXITCODE -ne 0) { throw 'Authenticode verification failed' }
    Write-Host 'Signed by:' (Get-AuthenticodeSignature "$out\MicroadsOffice.exe").SignerCertificate.Subject
  } else {
    Write-Host "NOTE: not Authenticode-signed (no -CodeSignThumbprint): browsers/SmartScreen will warn on direct downloads." -ForegroundColor Yellow
  }

  Copy-Item "$out\MicroadsOffice.exe" $setup -Force
  & $py -I "$b\scripts\signkey.py" sign $setup $version | Set-Content -Encoding ascii "$setup.sig.json"
  if ($LASTEXITCODE -ne 0) { throw 'Signing failed' }
  Write-Host "Setup  : $setup ($([math]::Round((Get-Item $setup).Length / 1MB, 1)) MB)"
  Write-Host "Sig    : $setup.sig.json"

  if ($Store) {
    Write-Host "`n== Microsoft Store package (MSIX) ==" -ForegroundColor Cyan
    if (-not $StoreIdentityName -or -not $StorePublisher) { throw '-Store needs -StoreIdentityName and -StorePublisher from Partner Center (Product identity)' }
    & "$root\tools\store\make_msix.ps1" -IdentityName $StoreIdentityName -Publisher $StorePublisher -PublisherDisplay $StorePublisherDisplay -Exe "$out\MicroadsOffice.exe"
  }
  } # end of -not $WebSetupOnly

  Write-Host "`n== Web installer (never changes) ==" -ForegroundColor Cyan
  dotnet build src\MicroadsOffice.WebSetup\MicroadsOffice.WebSetup.csproj -c Release
  if ($LASTEXITCODE -ne 0) { throw 'Web installer build failed' }
  $webBuilt = Join-Path $root 'src\MicroadsOffice.WebSetup\bin\Release\net462\MicroadsOffice_Setup.exe'
  $builtHash = (Get-FileHash $webBuilt -Algorithm SHA256).Hash.ToLowerInvariant()
  if (-not (Test-Path $webFrozen) -or $NewWebSetup) {
    if (Test-Path $webFrozen) { Write-Host 'NEW web installer requested (-NewWebSetup): its browser / SmartScreen reputation starts again from zero.' -ForegroundColor Yellow }
    New-Item -ItemType Directory -Force -Path (Split-Path $webFrozen) | Out-Null
    Copy-Item $webBuilt $webFrozen -Force
    Set-Content -Encoding ascii -Path "$webFrozen.sha256" -Value $builtHash
  }
  $frozenHash = (Get-FileHash $webFrozen -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($frozenHash -ne (Get-Content "$webFrozen.sha256" -Raw).Trim()) { throw "release\web-setup\MicroadsOffice_Setup.exe was modified (hash $frozenHash)" }
  if ($builtHash -eq $frozenHash) { Write-Host "Web installer: identical to the published one ($frozenHash)" }
  else { Write-Host "Web installer: source differs from the frozen binary — keeping the frozen one ($frozenHash). Use -NewWebSetup only for a necessary change." -ForegroundColor Yellow }

  if ($UiTests) {
    if (-not (Test-Path $setup)) { throw "Full setup $setup not found (build it first)" }
    Write-Host "`n== Web installer end-to-end test (UI Automation; windows appear for ~3 minutes) ==" -ForegroundColor Yellow
    dotnet build tests\WebSetup.E2E\WebSetup.E2E.csproj -c Release -o "$b\tools\e2e-web"
    if ($LASTEXITCODE -ne 0) { throw 'Web installer E2E build failed' }
    & "$b\tools\e2e-web\WebSetup.E2E.exe" $webFrozen $setup "$setup.sig.json" "$b\e2e-web\work" $py "$b\scripts\signkey.py" "$b\secrets\release-signing-backup.pem"
    if ($LASTEXITCODE -ne 0) { throw 'Web installer E2E failed (see .build\e2e-web\report-websetup.txt)' }
  }

  if ($UiTests -and -not $WebSetupOnly) {
    Write-Host "`n== PDF UI tests (mouse/keyboard are used: do not touch the PC for a few minutes) ==" -ForegroundColor Yellow
    $env:USERPROFILE = $origProfile; $env:APPDATA = $origAppData; $env:LOCALAPPDATA = $origLocal; Remove-Item Env:HOME -ErrorAction SilentlyContinue
    & "$root\tests\MicroadsOffice.Tests\bin\Release\net10.0-windows\MicroadsOffice.Tests.exe" ui "$b\ui-out" "$out\MicroadsOffice.exe"
  }

  if ($Publish) {
    Write-Host "`n== Publish to microads.vn/office ==" -ForegroundColor Cyan
    if (-not $WebSetupOnly) {
      $pargs = @("-I", "$root\server\deploy\publish_release.py", $setup, "$setup.sig.json")
      if ($NotesVi) { $pargs += @('--notes-vi', $NotesVi) }
      if ($NotesEn) { $pargs += @('--notes-en', $NotesEn) }
      if ($ForceUpdate) { $pargs += '--force' }
      & $py @pargs
      if ($LASTEXITCODE -ne 0) { throw 'Publishing failed' }
      & $py -I "$root\tools\winget\make_manifests.py" "$setup.sig.json"
      Write-Host "winget: submit the folder above to github.com/microsoft/winget-pkgs (see tools\winget\README.md)" -ForegroundColor Yellow
    }
    # The web installer is uploaded once; the server keeps the existing one when it is identical (and refuses a different one).
    $wargs = @("-I", "$root\server\deploy\publish_websetup.py", $webFrozen)
    if ($NewWebSetup) { $wargs += '--replace' }
    & $py @wargs
    if ($LASTEXITCODE -ne 0) { throw 'Publishing the web installer failed' }
  }
}
finally {
  dotnet build-server shutdown 2>$null | Out-Null
  Pop-Location
}
