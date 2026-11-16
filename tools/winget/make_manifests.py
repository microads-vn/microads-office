"""Writes the Windows Package Manager (winget) manifests of one Microads Office release.

usage: python -I tools/winget/make_manifests.py <MicroadsOffice_Setup_x.y.z.exe.sig.json> [out folder]
Output: tools/winget/manifests/m/Microads/Office/<version>/ (same layout as github.com/microsoft/winget-pkgs).
winget runs the full installer with:  --install --silent [--dir "<location>"]  (per-user install, no administrator rights).
"""
import json, os, sys

PACKAGE_ID = 'Microads.Office'
SCHEMA = '1.10.0'
sig = json.load(open(sys.argv[1], encoding='utf-8'))
ver = sig['version']
here = os.path.dirname(os.path.abspath(__file__))
out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, 'manifests', 'm', 'Microads', 'Office', ver)
os.makedirs(out, exist_ok=True)
url = f'https://microads.vn/office/download/MicroadsOffice_Setup_{ver}.exe'
exts = ['docx', 'docm', 'dotx', 'rtf', 'txt', 'xlsx', 'xlsm', 'xltx', 'csv', 'tsv', 'pdf']

def write(name, kind, body):
    text = f'# yaml-language-server: $schema=https://aka.ms/winget-manifest.{kind}.{SCHEMA}.schema.json\n\n' + body.strip() + '\n'
    with open(os.path.join(out, name), 'w', encoding='utf-8', newline='\n') as f: f.write(text)

write(f'{PACKAGE_ID}.yaml', 'version', f"""
PackageIdentifier: {PACKAGE_ID}
PackageVersion: {ver}
DefaultLocale: en-US
ManifestType: version
ManifestVersion: {SCHEMA}
""")

write(f'{PACKAGE_ID}.installer.yaml', 'installer', f"""
PackageIdentifier: {PACKAGE_ID}
PackageVersion: {ver}
Platform:
- Windows.Desktop
MinimumOSVersion: 10.0.17763.0
InstallerType: exe
Scope: user
InstallModes:
- interactive
- silent
- silentWithProgress
InstallerSwitches:
  Silent: --silent
  SilentWithProgress: --silent
  InstallLocation: --dir "<INSTALLPATH>"
  Custom: --install
UpgradeBehavior: install
ProductCode: MicroadsOffice
FileExtensions:
""" + ''.join(f'- {e}\n' for e in exts) + f"""AppsAndFeaturesEntries:
- DisplayName: Microads Office
  ProductCode: MicroadsOffice
Installers:
- Architecture: x64
  InstallerUrl: {url}
  InstallerSha256: {sig['sha256'].upper()}
ManifestType: installer
ManifestVersion: {SCHEMA}
""")

write(f'{PACKAGE_ID}.locale.en-US.yaml', 'defaultLocale', f"""
PackageIdentifier: {PACKAGE_ID}
PackageVersion: {ver}
PackageLocale: en-US
Publisher: Microads Technology Co., Ltd.
PublisherUrl: https://microads.vn
PublisherSupportUrl: https://microads.vn/office/
PrivacyUrl: https://microads.vn/office/privacy
Author: Microads Technology Co., Ltd.
PackageName: Microads Office
PackageUrl: https://microads.vn/office/
License: Freeware
Copyright: © Microads — microads.vn
ShortDescription: Free office suite with a word processor, spreadsheets and a PDF reader with digital signatures.
Description: |-
  Microads Office is a free office suite published by Microads Technology Co., Ltd. (Vietnam).
  It contains Word (documents, .docx), Excel (spreadsheets with formulas and charts, .xlsx/.csv) and PDF
  (read, digitally sign and print). Vietnamese and English user interface, no account needed.
  Updates are digitally signed by Microads and verified before they are installed.
Moniker: microads-office
Tags:
- docx
- excel
- office
- office-suite
- pdf
- pdf-reader
- spreadsheet
- vietnamese
- word
- xlsx
ReleaseNotesUrl: https://microads.vn/office/
ManifestType: defaultLocale
ManifestVersion: {SCHEMA}
""")

write(f'{PACKAGE_ID}.locale.vi-VN.yaml', 'locale', f"""
PackageIdentifier: {PACKAGE_ID}
PackageVersion: {ver}
PackageLocale: vi-VN
Publisher: Công ty TNHH Công nghệ Microads
PublisherUrl: https://microads.vn
PublisherSupportUrl: https://microads.vn/office/
PrivacyUrl: https://microads.vn/office/privacy
Author: Công ty TNHH Công nghệ Microads
PackageName: Microads Office
PackageUrl: https://microads.vn/office/
License: Miễn phí
ShortDescription: Bộ ứng dụng văn phòng miễn phí gồm soạn thảo văn bản, bảng tính và đọc, ký số PDF.
Description: |-
  Microads Office là bộ ứng dụng văn phòng miễn phí do Công ty TNHH Công nghệ Microads phát hành.
  Gồm Word (soạn thảo văn bản, .docx), Excel (bảng tính, công thức, biểu đồ, .xlsx/.csv) và PDF (đọc, ký số, in).
  Giao diện tiếng Việt và tiếng Anh, không cần tài khoản. Bản cập nhật có chữ ký số của Microads và được kiểm tra trước khi cài.
Tags:
- excel
- ky-so
- office
- pdf
- van-phong
- word
ManifestType: locale
ManifestVersion: {SCHEMA}
""")
print(out)
