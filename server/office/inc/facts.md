# Microads Office — product facts (the ONLY source for claims about the product)

Publisher: Microads Technology Company Limited (Công ty TNHH Công nghệ Microads), Ho Chi Minh City, Vietnam — https://microads.vn
Product page and download: https://microads.vn/office/ (Vietnamese) · https://microads.vn/office/en/ (English)
Price: free forever. No account, no sign-up, no subscription, no trial period, no feature limits. Open source under GPL-3.0
(source code: https://github.com/microads-vn/microads-office).
Platform: Windows 10 and Windows 11, 64-bit ONLY. There is NO macOS, Linux, Android, iOS or web version — never claim one.
Interface languages: Vietnamese and English (switch with the flag button).
Installation: one installer; per-user install without administrator rights; small web installer (80 KB) that downloads the latest
signed version; full offline installer (~55 MB); silent install for IT (`--silent`). Microsoft Store release is pending review.
Updates: built-in "Check for updates"; every update is verified with the Microads ECDSA release signature before installing.
Publisher introduction: the apps show a short "Microads Office — published by Microads" window with the company's services and a
link to microads.vn at start-up and when a file is opened from a folder; it closes immediately with one click (no waiting). There are
no third-party ads, no ads inside documents and no watermark on exported PDFs. Describe it honestly as a publisher introduction window,
never as "no ads at all".
Privacy: documents never leave the computer. Only automatic connection = update check (can be switched off).
Works offline (except the optional AI assistant, which needs Internet or a local Ollama server).

## Microads Office Word (word processor)
- Opens/saves .docx (Office Open XML); opens .rtf, .txt; saves .txt, .html, .rtf. Export to PDF with selectable text. Print with preview.
- CANNOT open old binary .doc (Word 97-2003) or .odt — it asks the user to re-save as .docx in another program.
- Page-accurate print layout, ruler, zoom, navigation pane by headings, full-screen reading mode.
- Fonts: bold/italic/underline (single/double/dotted)/strikethrough/subscript/superscript, change case, colours, highlight, clear formatting.
- Paragraphs: bullets, numbering, multilevel lists, indents, first-line/hanging indent, alignment incl. justify, line spacing
  (1.0/1.15/1.5/2.0/exact), spacing before/after, shading, borders, show/hide ¶, sort paragraphs A→Z.
- Styles: Normal, Heading 1–3, Title, Subtitle, Quote, No Spacing; basic custom styles.
- Find & Replace (match case, whole word), Go to page, word count (words, characters, paragraphs, pages).
- Insert: page break (Ctrl+Enter), tables (grid picker or dialog), pictures inline (PNG/JPG/BMP/GIF, paste from clipboard, resize
  handles), hyperlinks (Ctrl+K), header/footer, page numbers ("Page X of Y"), date/time, special symbols, horizontal line.
- NOT supported: text wrapping around pictures (pictures are inline only), shapes, SmartArt, charts inside Word, text boxes, WordArt,
  equations, footnotes/citations/index, comments, Track Changes, document compare, newspaper columns, line numbers, multiple sections
  with different paper sizes, drag-and-drop of selected text, macros/VBA, real-time co-authoring, OneDrive/SharePoint.
- Layout: margins incl. a preset for Vietnamese administrative documents (top 20, bottom 20, left 30, right 15 mm), portrait/landscape,
  paper A4/A5/A3/Letter/Legal/B5, page colour, text watermark (DRAFT, CONFIDENTIAL…).
- Automatic table of contents from Heading 1–3 with update.
- Tables: insert/delete rows and columns, merge cells, cell shading, borders, Tab moves between cells / adds a row.
- Templates: blank, Vietnamese administrative letter (Decree 30/2020/ND-CP), letter, CV/résumé, report.
- Spell check: uses the Windows spell checker for installed languages (basic).
- Vietnamese typing with Unikey, EVKey or the Windows Telex keyboard.
- Standard Word shortcuts: Ctrl+B/I/U, Ctrl+L/E/R/J, Ctrl+]/[, Ctrl+Shift+=, Ctrl+=, Ctrl+F/H/G, Ctrl+P, Ctrl+S, Ctrl+Z/Y.
- Paste: screenshots, pictures, files from File Explorer (pictures inserted, other files as links), rich text from browsers.
- AutoRecover backup and recovery after a crash.

## Microads Office Excel (spreadsheet)
- Opens/saves .xlsx (values, formulas, styles, merged cells, column widths/row heights, freeze panes, multiple sheets, defined names,
  conditional formatting, data validation, notes, hyperlinks, basic charts); .csv/.txt (UTF-8). Export PDF, save HTML.
- CANNOT open .xls (Excel 97-2003) or .ods.
- 1,048,576 rows × 16,384 columns; formula bar and Name Box; fill handle (series, dates, weekdays, months, formulas);
  function-name suggestions while typing; Paste Special (values, formulas, formats, transpose); hide/unhide; sheet tabs with colours.
- Formatting: fonts, fill, borders, alignment, wrap text, merge, number formats (general, number, currency ₫/$, accounting, dates,
  time, percent, fraction, scientific, text, custom codes), conditional formatting (greater/less/between/equal/text contains/duplicates/
  top 10/data bars/colour scales), Format as Table, cell styles.
- Data: sort A→Z/Z→A and multi-level custom sort, AutoFilter, remove duplicates, text to columns, data validation (drop-down lists,
  numbers, dates, length), sheet protection with password, cell notes.
- Charts: column, bar, line, pie, doughnut, area, scatter, with title and legend.
- Page layout: margins, orientation, paper size, print area, repeat title rows, fit to one page wide, print gridlines/headings.
- Formulas: Insert Function dialog by category, Name Manager, Show Formulas, automatic/manual calculation, F9.
- 502 worksheet functions covering all Excel 365 categories: math & trig, statistical, logical, lookup & reference (VLOOKUP, HLOOKUP,
  XLOOKUP, INDEX, MATCH, XMATCH, OFFSET, INDIRECT…), text (TEXTJOIN, TEXTSPLIT, REGEX…), date & time (DATEDIF, NETWORKDAYS, WORKDAY,
  EOMONTH…), financial (PMT, NPV, IRR, XIRR, bonds…), engineering, database (DSUM…), information, web (ENCODEURL, FILTERXML).
- Dynamic arrays that spill: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS, CHOOSEROWS,
  TOCOL, TOROW, WRAPROWS, WRAPCOLS, EXPAND; #SPILL! error; spill reference A1#.
- LET, LAMBDA (named, recursive, optional [x] parameters), MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, GROUPBY, PIVOTBY.
- Vietnamese function DOCSO(number, [unit]): writes a number in Vietnamese words, already capitalised and ending with the unit
  "đồng" by default — DOCSO(120000) = "Một trăm hai mươi nghìn đồng", DOCSO(1005) = "Một nghìn không trăm linh năm đồng",
  DOCSO(-5) = "Âm năm đồng"; DOCSO(35,"") = "Ba mươi lăm" (no unit); another unit text can be given as the 2nd argument.
  DOCSO exists only in Microads Office Excel (Microsoft Excel has no such function).
- Files saved with the new functions open correctly in Microsoft Excel 365.
- NOT supported: PivotTable (planned), sparklines, slicers, 3D maps, Power Query, database connections, What-If, Solver, formula
  auditing arrows, Watch Window, Evaluate Formula, macros/VBA, cube functions, online-only functions (WEBSERVICE, STOCKHISTORY, IMAGE,
  TRANSLATE, COPILOT, PY, FORECAST.ETS), text rotation.

## Microads Office PDF
- Fast PDF viewer (PDFium engine), print.
- Digital signatures: sign PDF with a USB token (the token's driver must be installed so its certificate appears in the Windows
  certificate store — e.g. Vietnamese digital-signature tokens such as those used for e-tax) or a .pfx certificate file; PAdES / PKCS#7 signatures; verify existing signatures; image stamps (e.g. a scanned signature or company seal);
  paste or drag a picture onto a page to place it as a stamp.

## AI assistant (optional, in Word, Excel and PDF)
- Purple sparkle button on the Home tab opens a side pane. Drafts, rewrites, translates and proofreads straight into the document
  (one Ctrl+Z undoes it); writes formulas, tables and charts into the sheet; explains formulas; summarises and answers questions about PDFs.
- Connect with: "Continue with ChatGPT" (the user's own ChatGPT Plus/Pro plan), the user's own Claude or OpenAI API key, an
  OpenAI-compatible server such as Ollama running locally (free, fully offline), or the user's Claude Pro/Max plan through Claude Desktop
  (built-in MCP server).
- Microads does not provide or sell AI credits. The assistant connects only when the user clicks Send, directly to the chosen service,
  never through Microads servers. Keys are encrypted on the computer (Windows DPAPI).

## Honest comparison notes
- Microads Office is not affiliated with Microsoft. "Microsoft", "Word", "Excel", "Office 365" are Microsoft trademarks; mention them
  only to compare or to explain compatibility.
- Good fit: everyday documents, administrative letters, spreadsheets with formulas, signing PDFs, offline use, no subscription.
- Not a fit: heavy collaboration, Track Changes reviews, macros/VBA, PivotTables, Mac/Linux users.
