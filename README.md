# Microads Office

**Free office suite for Windows: Word (documents), Excel (spreadsheets) and PDF (read, digitally sign, print) — in one installer.**
Published by **Microads Technology Company Limited** (Công ty TNHH Công nghệ Microads), Ho Chi Minh City, Vietnam — [microads.vn](https://microads.vn).
Always free, no account, Vietnamese and English user interface.

[Download](https://microads.vn/office/) · [Code signing policy](CODE_SIGNING_POLICY.md) · [Privacy policy](https://microads.vn/office/privacy) · [License: GPL-3.0](LICENSE)

## Features

| App | What it does |
|---|---|
| **Microads Office Word** | Page-accurate word processor: styles, tables, images, headers/footers, table of contents, Vietnamese administrative document templates (Decree 30/2020), `.docx` / `.rtf` / `.txt` read & write, export to PDF, print |
| **Microads Office Excel** | Spreadsheets with ~230 functions (incl. `DOCSO` — numbers to Vietnamese words), charts, conditional formatting, sort & filter, `.xlsx` / `.csv` read & write, print / PDF |
| **Microads Office PDF** | Fast PDF viewer (PDFium), digital signatures (PAdES / PKCS#7 with USB tokens or `.pfx`), signature verification, image stamps, print |
| **AI assistant** (Word, Excel, PDF) | Side pane that drafts, rewrites, translates and proofreads straight into the document (one undo step), writes formulas / tables / charts into the sheet, summarises PDFs. Connect with **Continue with ChatGPT** (your ChatGPT Plus/Pro plan), a Claude or OpenAI API key, an OpenAI-compatible server (Ollama), or use your Claude Pro/Max plan from **Claude Desktop** through the built-in MCP server (`MicroadsOffice.exe --mcp`) |
| **Hub & installer** | One self-contained executable; per-user installation without administrator rights; file associations; uninstaller; automatic updates verified with the Microads ECDSA release signature |

Full feature table: [docs/BANG_TINH_NANG.md](docs/BANG_TINH_NANG.md) (Vietnamese). User guide: [docs/HUONG_DAN_SU_DUNG.md](docs/HUONG_DAN_SU_DUNG.md).

## Download and install

* **Web installer** (recommended): <https://microads.vn/office/download/MicroadsOffice_Setup.exe> — a small program that downloads the latest
  release from microads.vn, verifies its signature and starts the setup.
* **PowerShell** (no browser): `irm https://microads.vn/office/install.ps1 | iex`
* Silent install for IT: `MicroadsOffice_Setup.exe --silent [--dir "D:\Apps\Microads Office"] ` (add `--no-auto-update` to switch off automatic update checks).

Requirements: Windows 10 or 11, 64-bit.

## Privacy

Documents never leave your computer. The only connection the apps make on their own is the **update check** (it sends the installed version
number to `https://microads.vn/office/api/manifest.json`). It can be switched off during installation or later in *Help › Settings*.
The optional **AI assistant** connects only when you click *Send*, and only to the AI service you chose (your ChatGPT account, your own
Claude / OpenAI API key, your own server such as Ollama, or Claude Desktop) — never through Microads servers. Keys and sign-ins are
encrypted with Windows DPAPI and stay on your computer.
Details: <https://microads.vn/office/privacy>.

## Building from source

Requirements: Windows 10/11 x64, [.NET SDK 10](https://dotnet.microsoft.com/download).

```powershell
dotnet build MicroadsOffice.sln -c Release
dotnet publish src\MicroadsOffice -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=false -p:EnableCompressionInSingleFile=true -p:DebugType=none -o out
```

`build.ps1` is the release script used by Microads (tests, packaging, release signature, publishing). Release builds that are distributed are
built by GitHub Actions ([.github/workflows](.github/workflows)) and Authenticode-signed through SignPath (see the code signing policy).

| Folder | Contents |
|---|---|
| `src/MicroadsOffice.Core` | Shared engine: PDF parsing, PDFium, digital signatures, localisation, settings, update service |
| `src/MicroadsOffice.UI` | Shared WinForms UI (ribbon, backstage, start screen, dialogs) |
| `src/MicroadsOffice.Word`, `.Excel`, `.Pdf` | The three applications |
| `src/MicroadsOffice` | The single executable: dispatcher, hub, installer / uninstaller |
| `src/MicroadsOffice.WebSetup` | The small web installer (.NET Framework 4.6.2) |
| `tests/` | Automated tests (core, Word, Excel, web installer end-to-end) |
| `server/` | The download & update site microads.vn/office (PHP) |
| `docs/` | Architecture, feature table, user guide, release notes (Vietnamese) |

## License

Copyright © 2026 Microads Technology Company Limited.

Microads Office is free software: you can redistribute it and/or modify it under the terms of the **GNU General Public License version 3**
as published by the Free Software Foundation ([LICENSE](LICENSE)). It is distributed WITHOUT ANY WARRANTY.

Third-party components and their licenses: [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

**Trademarks:** "Microads", "Microads Office" and the Microads logo are trademarks of Microads Technology Company Limited. The GPL does not
grant permission to use them; modified versions must be distributed under a different name and without the Microads logo.

---

## Tiếng Việt

**Microads Office** là bộ ứng dụng văn phòng miễn phí (Word, Excel, PDF) do **Công ty TNHH Công nghệ Microads** phát hành, mã nguồn mở theo
giấy phép **GPL-3.0**. Tải về tại <https://microads.vn/office/>. Tài liệu tiếng Việt nằm trong thư mục [docs](docs).
