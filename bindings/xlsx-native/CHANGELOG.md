# @betteroffice/xlsx-native

## 0.6.2

### Patch Changes

- 305d381: A font picked for Chinese, Japanese or Korean text in a document now stays after saving and reopening in Word, and Latin text keeps its East Asian font. Spreadsheet cells given a font no longer follow the theme font in Excel.
- 8475b31: The Linux GNU binaries load again on glibc 2.17 and newer on both supported architectures.
- a42798b: Documents saved on Windows are now byte-identical to the same documents saved on Linux and macOS. Saved content does not change.
- 372e1e3: Printed sheets keep their stored row heights when the default row height is below the font's natural height, so they paginate closer to Excel.
- Updated dependencies [e163044]
- Updated dependencies [662a971]
  - @betteroffice/fonts@0.5.1
  - @betteroffice/fonts-cjk@0.5.1
  - @betteroffice/fonts-intl@0.5.1
  - @betteroffice/fonts-latin@0.5.1

## 0.6.1

### Patch Changes

- 365ae9d: The Linux native packages load again on Ubuntu 24.04, Debian 12 and other distributions with glibc 2.35 or newer.

## 0.6.0

### Minor Changes

- cf72bb4: The native packages now export the full `@betteroffice/docx`, `@betteroffice/xlsx` and `@betteroffice/pptx` APIs as ES modules, and `require()` of them needs Node 20.19 or 22.12: `engines` now says so. `@betteroffice/agents`, which depends on them, requires Node 22.12.
- f1e7906: Adds experimental `fontPreflight` and `preloadPreflightWasm` for the active slide and the opening sheet, without opening the editing engine. DOCX `fontPreflight` answers now follow the engine for notes, the main part, numbering, theme script fonts, style defaults and table columns.
- 702b174: The calculation engine now evaluates RATE, IRR, MIRR, XNPV, XIRR, IPMT, PPMT, CUMIPMT, CUMPRINC, ISPMT, DB, DDB, VDB, SLN, SYD, EFFECT, NOMINAL, PDURATION, RRI, FVSCHEDULE, DOLLARDE, DOLLARFR and CELL. RATE follows Excel's own iteration, so results match Excel's saved values.
- dac1245: The calculation engine now evaluates GAMMA, GAMMALN and the gamma, beta, normal, lognormal, chi-square, t and F distribution functions with their inverses and compatibility names, ERF and ERFC, PHI, GAUSS, CHISQ.TEST, T.TEST, F.TEST and the CONFIDENCE functions.
- 317d85c: The calculation engine now evaluates Excel's database functions DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR and DVARP, including OR rows, repeated columns, wildcards and computed criteria.
- f92d560: The calculation engine now evaluates the binomial, negative binomial, hypergeometric, Poisson, exponential and Weibull distribution functions, FISHER, STANDARDIZE, PERMUT, PEARSON, STEYX, SKEW.P, the A-variance functions, the exclusive percentile and rank functions, MODE.MULT, PROB, Z.TEST, GROWTH and LOGEST.
- c06b68d: The calculation engine now evaluates Excel's engineering functions: base conversions (BIN2DEC through OCT2HEX), the BIT functions, DELTA, GESTEP, COMPLEX, every IM complex-number function and the four Bessel functions. Complex results are written as Excel writes them.
- 451d540: The calculation engine now evaluates GROUPBY, PIVOTBY, TRIMRANGE and EUROCONVERT, with function names such as SUM as arguments. PRODUCT of a block without numbers returns 0 and ARRAYTOTEXT keeps 15 significant digits. GETPIVOTDATA, FORECAST.ETS and two GROUPBY/PIVOTBY function layouts keep Excel's saved values.
- 74407d2: The calculation engine now evaluates more information and text functions: ERROR.TYPE, INFO, ISFORMULA, ISNONTEXT, ISOMITTED, ISREF, SHEET, SHEETS, TYPE, UNICHAR, UNICODE, DOLLAR, VALUETOTEXT, ARRAYTOTEXT, ASC, JIS, DBCS, BAHTTEXT, PHONETIC, ENCODEURL and the byte-count text functions such as LENB.
- e5bb91a: The calculation engine now evaluates 33 more math and trig functions, including the hyperbolic and reciprocal trig functions, ROMAN, ARABIC, BASE, DECIMAL, CEILING.MATH, FLOOR.MATH, MDETERM, MINVERSE, MUNIT, RANDARRAY, the SUMX2 family, PERCENTOF and DAYS360.
- 2dbf721: The calculation engine now evaluates Excel's securities functions: the COUP functions, ACCRINT, ACCRINTM, PRICE, YIELD and their discount and at-maturity variants, DISC, INTRATE, RECEIVED, DURATION, MDURATION, the TBILL functions, the odd-period functions and AMORDEGRC and AMORLINC.
- 57a2246: Selecting a chart now shows eight resize handles. Dragging one resizes a grid-anchored, two-cell chart and commits the new size in a single undo step. Pinned and one-cell charts remain non-resizable.
- 4955aa3: Conditional formats are decided once per edit, at most 10 ms a frame after rule compilation and range statistics, so frames can run over. Rule formula names parse once per frame. Display lists gain `conditionalPending`, cells left undecided; the editor asks for frames until 0.
- 0d9427d: A number format a workbook declares for a built-in id is shown as declared, as in Excel. Shared states of a workbook that declares such an id use schema version 8, which clients on an older version do not open; all other workbooks are unchanged.
- 4520944: Adds `install_panic_hook`, matching the WebAssembly engine. The package installs it on load, so on macOS and Linux a Rust panic writes one `[xlsx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
- 3adc24b: `renderPng` draws in the native engine through the new raw `renderPngTarget(targetJson, optionsJson)`, which honours scale and maxWidth and maxHeight for viewports and ranges. The raw `renderPng(viewportJson)` and `renderRangePng(args)` are removed.
- 7a91087: Now exposes exactly the `@betteroffice/xlsx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openWorkbook(Buffer)`/`Workbook` facade is removed.
- ed7c58f: Adds the `openWorkbook` option `recalculate: 'stale'`, which shows the formula results a file cached and recalculates only when a formula has none or `calcPr` requests it. Defaults to `'all'`.
- 83d0750: Listed spreadsheet proposals now name their moved cells in `staleTargets` (`stale_targets` in Python), so hosts can warn before accepting. Python also gains `list_proposals()`, an alias of `proposals()`.
- e76f70d: Editors and direct writes now refuse locked cells, formats and chart moves and resizes on protected sheets by default, like Excel. Direct writes accept opt-in `protectedSheets: "allow"` (Python `protected_sheets="allow"`), which writes through and keeps the protection.
- d83fa7a: Renaming a sheet also rewrites data-validation, extension conditional-format and sparkline formulas that reference it, as Excel does; a rename they cannot follow is refused. Adds save-warning code `unrewritable-formula`.
- 2c28650: After `WorkbookHandle.recalculate()`, saved files carry the recalculated values; formulas the engine cannot evaluate keep the result the file stored. A save without edits, showing the stored active sheet, does not write the recalculation at open; applying a collaborative update counts as an edit.
- 0852138: Adds `saveWithWarnings()` (`save_with_warnings()` in Python): the saved bytes plus `{ code, part, element?, sheet?, message }` for each part a save kept as the file wrote it. Plain `save()` reports them as Node or Python warnings; ordinary saves carry none.
- f0ec30f: Tables stay on their own sheets when a sheet is inserted or deleted, so structured references keep reading the right cells; deleting a sheet drops its tables and turns references to them into `#REF!`. Undo restores them via a new internal `setTables` op.
- c1525f7: Formulas using functions that cannot be computed offline (CUBE functions, WEBSERVICE, RTD, STOCKHISTORY, IMAGE, add-in calls and similar), and formulas reading them, keep the value Excel saved, even inside IFERROR, and the calculation report lists them as `unevaluatedCells` in JS and `unevaluated` in Python.
- 4870194: Breaking: `DrawCmd` gains `Image`, `build_display_list_with_geometry` and `build_print_display_list_with_charts` take pictures, and `WorkbookHandle` requires `mediaBytes`. Screen and print display lists carry worksheet pictures, drawn by the native PNG export and by `paintDisplayList` given `resolveImage`; the JavaScript PNG export, bundled editor and viewer do not paint them yet.

### Patch Changes

- 57f4975: Arabic text set in Simplified Arabic or Traditional Arabic now falls back to the bundled Naskh face, closer to those fonts' style and width.
- fe023f2: Negative luminance offsets now darken DrawingML colours.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- b8395eb: `fontPackDiagnostics()` names installed font add-ons that cannot be used (another release, missing files, unreadable manifest). Native packages also warn once per pack with code `BETTEROFFICE_FONT_PACK`. Exports keep substituting those faces.
- b0d5d34: The native xlsx engine has the same 4 GiB memory budget per thread as the WebAssembly engine. A call past the budget fails the worker session with the typed out-of-memory error, and a failed workbook handle refuses later calls except dispose().
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 7f439b2: The native packages and each of their platform packages now ship the project's `LICENSE` file in the tarball, as the WebAssembly packages do. They declared Apache-2.0 without carrying the text.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- 4687b6f: `require()` of the native packages from CommonJS resolves, as it does for the WebAssembly packages: every export gains a `default` condition beside `import`. Before, it failed with `ERR_PACKAGE_PATH_NOT_EXPORTED`. Loading an ES module by `require()` needs Node 20.19+ or 22.12+.
- a121ec6: Native engine classes now have `[Symbol.dispose]`, the same function as `free()`, as on the wasm packages, so `using` declarations release them at scope exit.
- 9e9309a: Each worker loads as one file.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- d56440d: In Safari and WebKit, apps that bundle the editors with ES module workers (Vite `worker.format: 'es'`) could run a worker twice and fail on its first command. Each worker entry now loads its code once.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- fd25d97: Text in the Symbol font draws the Greek letters, arrows and math signs Office shows for it, in documents, slides and cells alike. Wingdings and Webdings run text uses the same shared slot table as bullets. Saved text is unchanged. Handle `SymbolFont::Symbol` in exhaustive matches.
- f77d138: Date functions in workbooks that use the 1904 date system now return 1904 serials: `DATE(2026,11,1)` is 44865, as in Excel. `nowSerial` is always a 1900-system serial, so TODAY(), NOW() and typed dates without a year match the calendar in 1904 workbooks too.
- f25a90c: Calculation matches Excel at numeric limits: numbers become text with 15 significant digits (`""&(0.1+0.2)` is "0.3"), formula literals keep 15 digits within Excel's range, subnormal results are 0, and MOD, MULTINOMIAL, BESSELK, the TBILL, T and F functions follow Excel's limits.
- e028c97: Full recalculation of an xlsx workbook can be faster and use less memory when its formulas, names, tables and sheets did not change since the previous calculation.
- f97d718: Recalculation is faster for formulas that depend on a sheet's last used row, such as SUMPRODUCT over whole columns.
- cc8855a: Full recalculation of an xlsx workbook is faster when its formulas did not change since the previous calculation.
- 3dd06dd: Browser `renderPng` without `fonts` paints with the bundled faces native uses, so canvas and native PNGs match; a face that cannot load rejects with `font-unavailable`. Native chart text in a family list (`Calibri, sans-serif`) now resolves its first family.
- b73727f: Inserting or deleting rows or columns inside a range a chart plots through a numeric cache whose points carry their own number format now rebuilds the cache, with each point's format taken from its cell, instead of being refused.
- 7aa4a92: XLSX chart colours apply their `lumMod`, `lumOff`, `tint`, `shade` and other colour transforms, so lighter or darker series shades and grey axis text draw as Excel shows them. `sysClr` and `prstClr` colours resolve, and translucent colours keep their alpha.
- 43f6296: A workbook whose chart reference holds a character XML cannot carry now opens, and its chart is kept as authored. An edit that would have to rewrite such a reference is refused up front, so saving never fails.
- 4990e8d: Native XLSX rendering draws charts much faster: a chart sheet that took about 400 ms now takes about 20 ms. Pixel output is unchanged.
- fec61ca: Row and column edits that reshape a range plotted through an unrebuildable cache or would split two sheets' shared chart part are refused upfront. Previously the next save failed. Cell edits under shifted ranges save. Drop the final original-workbook and sheet tuple from `chart_bytes`.
- e05c8d9: A chart that reads another workbook, carries literal series text, holds a lost `Sheet!#REF!` area or an extension range no longer blocks deleting or editing sheets it does not plot. Refusals in an untrusted package name a chart or pivot part.
- b5d9f51: Collaborators can add and remove conditional formats and set or clear auto filters, with undo. The schema is unchanged, but older peers refuse a room once a rule or filter is edited, so collaborating peers must update together.
- 34b0f7c: In a collaborative session, charts can be moved after merging cells or adding a hyperlink. The move used to fail with "remote update changes the frozen workbook structure".
- 7183550: Undoing a freeze pane that won a concurrent edit no longer fails: every peer shows the source file's pane and redo brings the winner back. Peers on older versions refuse that undo's update, so update them together.
- 8017c8a: Collaborators can set and delete defined names, with undo. The schema is unchanged, but older peers refuse a room once a name is edited, so collaborating peers must update together. Over-limit name tables save, report `overLimit`, but won't reopen here. `DefinedNamesResult` literals must set `over_limit`.
- 9ce4911: Collaborators can add sheets at the end of a workbook and merge cells in the editor. Sheets added at once under one name read `Name (2)`. Collaborating peers must update together: older peers refuse the new schema.
- 4955aa3: Conditional formats paint fill, font color, bold and italic, on screen and in PNG renders, only where the result is certain; after a cell edit, value paste or recalculation, an undecided cell keeps its previous format. On by default. Rust `xlsx_render` gains `CellFormats` and `build_display_list_with_geometry_and_formats`.
- 4b01734: Editing many defined names in one `applyOps` call now takes time and memory linear in the number of names. Re-committing all 66,000 names of a large workbook used to exhaust memory.
- 3a4bf3e: Deleting a sheet turns references to it into `#REF!` in formulas, defined names, conditional formats and data validations, as Excel saves them; 3-D spans narrow. Chart spans no longer change when a sheet inside them goes. Undo restores the text.
- a04beaf: A drawing that a sheet names with extra leading slashes, or through several relationships, is read once, so its charts and drawing objects are no longer listed several times, and such a workbook no longer fails to open with too many charts.
- 4955aa3: A conditional format whose fill names two different colors paints no fill, and a rule the calc engine cannot evaluate in full paints nothing. A cell with rich text runs takes only the fill. Rust: `DifferentialFormat` gains `is_released_reading_of`, `Stylesheet` gains `dxf_fill_colors_differ`, and `xlsx_calc::EvalContext::has_unhandled_unsupported_function` is public.
- 763916f: Number format codes in styles and chart caches are read and written as `ST_Xstring`, as Excel does: its `_x000a_` line break reads as a line break, any format set through the API saves as valid XML, and a literal `_xHHHH_` round-trips.
- 14b82ff: Shown and returned formula text drops Excel's storage prefixes (`_xlfn.`, `_xlws.`, `_xlpm.`, `_xleta.`): the formula bar, `readCells`, structured export, defined names, conditional formats, agent reads and Python `formula()` give `=SEQUENCE(2,2)`. Saved files keep the prefixes; re-committing a shown formula leaves the cell unchanged.
- 9d433be: Formulas written in BetterOffice that use newer Excel functions (XLOOKUP, IFS, LET, LAMBDA, FILTER and others) are saved with the `_xlfn.` prefixes Excel expects, so Excel no longer opens them as #NAME?. Unedited formulas keep their exact text.
- 7b5ec3f: Sheet, defined and font names are read and written as `ST_Xstring`, as Excel does: an escape Excel wrote reads decoded, and any name set through the API, including control characters, saves as valid XML and reads back unchanged.
- 6d749ac: PNG export uses the installed Latin, international and CJK font packs by default, reading faces from disk with nothing fetched.
- db2fca1: `renderPng` now paints with the faces passed in `fonts`, as the wasm engine does, and rejects with `font-unavailable` when a supplied face cannot be used. The option was previously ignored.
- 0036740: `renderPng` with a bad range now rejects with the same `Error` and `bad range: ...` message on native as on wasm, not a plain string. A1 ranges are read in either case on every engine, as Excel does: `a1:b2` is `A1:B2`.
- 445f40f: A native XLSX callback that uses or frees a workbook or snapshot builder already in use now gets an error, as with the wasm binding.
- 607aa41: Native XLSX rendering shapes and draws cell text through the shared font store and glyph cache, in about half the time. A no-break hyphen in a font without one draws as a hyphen, not a missing-glyph box. `FontStore` gains `register_shared`.
- 642b9b6: A save with no edit since open returns the file unchanged, also after an open that recalculated a full-calculation file, while the stored active sheet is shown and the archive is reusable. Applying a collaborative update or `recalculate()` counts as an edit.
- f955a24: Number formats with commas after the decimal digits, such as `0.0,,"M"` and `#,##0.0,"K"`, now divide by 1,000 per comma in cell display and `TEXT()`, as Excel does. A comma after a literal or percent sign no longer scales.
- a3630d1: Renaming a sheet rewrites the conditional-format formulas that reference it, as Excel does. Rules from the file keep everything else they hold, undo restores them, and rules added in the session no longer block the rename.
- 60abdd1: XLSX wrapped text shows only the lines whose top half fits its whole row or merged area, at any zoom and scroll position; the display list's `TextWrap` gains `maxLines`, which every painter obeys.
- 6dfe1b3: Renaming a sheet to a name holding a character XML cannot carry is refused as an invalid step when a formula, defined name or hyperlink references the sheet, instead of writing invalid XML; unreferenced sheets still take any name.
- e519913: Sheets whose first sheet view sets `showGridLines="0"` render on screen without gridlines and keep the setting on save. Print gridlines still follow the `gridlines` argument. Rust `SheetFormat` gains `hide_gridlines`.
- e200f5e: Speeds up style patches and number format changes on bounded ranges in standalone workbooks. Enabled by default, with the same saved files, updates and undo behavior.
- 9113fc7: A table or comments part that a sheet relates several ways is read once, so tables are not listed twice, comment counts are not multiplied, and such a workbook no longer fails to open with too many tables.
- be0f1ae: A table cell whose own font has no color, in a workbook whose normal font has one, keeps its automatic text color instead of the table style's, as Excel prints it.
- 06ef552: Excel tables now render their built-in or workbook-defined table style by default: header, total, banded and first or last column fills, font colours and bold. Breaking: Rust `Table` gains a public `style` field, so struct literals must set it.
- 169bdf2: XLSX rich text cells paint each run in its own size, style, colour and available font on screen, in PNG and print. Formula, number-formatted, rotated, stacked, right-to-left cells and cells without room stay plain. Rust code building or fully matching `DrawCmd::Text` must add `runs`.
- acf6184: Cell fonts with `<u val="single"/>`, `double` or accounting underlines now read as underlined, and double kinds paint two lines. Accounting underlines draw as plain single or double for now. Text commands add `underlineStyle`, captured fonts add `underline_style`; Rust `Font.underline` is now a `RunUnderline`.
- 5ee5cd3: Formulas calling a function the engine does not know keep the value Excel saved, even inside IFERROR, and are listed as unevaluated with the function name. GAMMA, GAMMALN and FACT are exact at integers.
- 14251f8: Line breaks stored as CR LF or a lone CR read as one line feed, as in Excel, so table references to multi-line column names calculate instead of returning #REF!. Saves keep carriage returns in cell text, formulas and hyperlink tooltips.
- 916cbe2: Cell text decodes Excel's `_xHHHH_` escapes, so `_x000D_` reads as a carriage return, and saves write control characters the way Excel does. Opened formulas keep control characters; new ones store them as `CHAR(n)`, and chart caches over such cells save.
- Updated dependencies [f99e53a]
- Updated dependencies [de8f092]
- Updated dependencies [1a82d79]
- Updated dependencies [3609ce4]
- Updated dependencies [5999962]
- Updated dependencies [57f4975]
- Updated dependencies [d60b8e4]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [f52a44a]
- Updated dependencies [b8395eb]
- Updated dependencies [d60b8e4]
- Updated dependencies [53502c6]
- Updated dependencies [58a8ca6]
- Updated dependencies [ec04baa]
- Updated dependencies [d60b8e4]
- Updated dependencies [21a60e5]
- Updated dependencies [6aefa34]
- Updated dependencies [19c20c3]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [18b3628]
- Updated dependencies [dc78510]
- Updated dependencies [54d400a]
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
- Updated dependencies [9e9309a]
  - @betteroffice/fonts@0.5.0
  - @betteroffice/fonts-cjk@0.5.0
  - @betteroffice/fonts-intl@0.5.0
  - @betteroffice/fonts-latin@0.5.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
