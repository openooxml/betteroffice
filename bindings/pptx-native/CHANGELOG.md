# @betteroffice/pptx-native

## 0.6.0

### Minor Changes

- b2ea7f5: The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.
- a45ba76: `renderPng({ slide }, options?)` on the presentation handle renders through the native rasterizer, with the same targets, options and errors as the browser.
- f1e7906: Adds experimental `fontPreflight` and `preloadPreflightWasm` for the active slide and the opening sheet, without opening the editing engine. DOCX `fontPreflight` answers now follow the engine for notes, the main part, numbering, theme script fonts, style defaults and table columns.
- 9ac1ffe: Adds `configureFontCdn` to load missing font packs from a configured CDN. Installed packs take priority, and CDN loading remains off by default.

### Patch Changes

- fa2717d: Every engine now reads decimal numbers back exactly. Drawing groups whose scaled child coordinates have long decimals stay intact on save.
- 57f4975: Arabic text set in Simplified Arabic or Traditional Arabic now falls back to the bundled Naskh face, closer to those fonts' style and width.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- b8395eb: `fontPackDiagnostics()` names installed font add-ons that cannot be used (another release, missing files, unreadable manifest). Native packages also warn once per pack with code `BETTEROFFICE_FONT_PACK`. Exports keep substituting those faces.
- 13a0dc0: EMF and WMF pictures that draw the same bitmap several times, or whose EMF+ drawing falls back to the GDI one, no longer exceed the bitmap budget: equal bitmaps are decoded and counted once, and a set-aside pass's bitmaps are released first.
- 92ca93c: EMF and WMF pictures that draw the same bitmap several times now rasterize within the image budget in native PNG exports: a repeated bitmap's decoded pixels are counted once, as the renderer decodes and frees each draw on its own.
- 79ae853: Native rendering now draws JPEG and PNG pictures past the pixel cap at a reduced size, matching the canvas downscale, instead of leaving them out. Pictures whose reduced decode would still exceed the memory budget are refused as before.
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- a121ec6: Native engine classes now have `[Symbol.dispose]`, the same function as `free()`, as on the wasm packages, so `using` declarations release them at scope exit.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- 3418edd: `renderPng` in the browser lays slides out with the bundled faces the native engine uses, so line breaks and row heights match native. On both engines, a needed face that cannot load now rejects with `PngExportError` (`font-unavailable`) naming the face.
- cade718: Kanji in Japanese decks now load and paint with Noto Sans JP instead of also fetching Noto Sans SC: Han text takes its CJK face from the run, paragraph mark or style language (Korean, Traditional and Simplified Chinese likewise).
- 0530412: The same font bytes registered under several family names are now held once and no longer count against the 256-face limit. Opening a deck with more than 256 distinct faces now fails with `PptxResourceLimitError` instead of a closed worker connection.
- e664196: The PowerPoint engine is about 3 KB smaller compressed: the slide layout cache compares slide floats with a direct walk instead of a generic serializer. Layout and cache behaviour are unchanged.
- 7daa244: Adds an optional `data` to chart display-list primitives with categories and each series' values as its data labels format them. Screen reader slide mirrors read a chart as a figure holding its values as a table; charts without data keep their summary.
- 3072028: Screen reader slide mirrors describe a chart with no cached values as having no data.
- d675bea: Adds paragraph `links` to slide display lists. Screen reader slide mirrors expose safe external links and slide navigation as keyboard accessible links by default.
- f73c684: Adds an optional `grid` to table display-list primitives with rows, columns, spans and header cells. Screen reader slide mirrors read tables as rows of header and data cells; older engines keep the flat reading.
- 8d0046d: Slides with a CMYK TIFF picture that has an alpha channel now export on the native engine instead of failing with `picture-unavailable`, using the same decoder the canvas engine uses.
- 7f42859: `renderPng` works with no fonts passed: it uses the deck's embedded faces and the installed font packs, and honours `options.fonts`. `layoutSlide` results are unchanged by an export.
- ec04baa: `@betteroffice/pptx-native` now installs `@betteroffice/fonts-latin`, `@betteroffice/fonts-intl` and `@betteroffice/fonts-cjk` and reads their faces from disk on demand, so Latin, international and CJK text renders with real faces without any fetch. On by default.
- 7f6277c: Native pptx slide PNGs reject with `picture-unavailable` when a picture cannot be drawn, as native docx pages do, instead of leaving it out. A picture frame with no image still renders.
- db4ab57: Now exposes exactly the `@betteroffice/pptx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openPresentation(Buffer)`/`Presentation` facade is removed.
- 1676610: Pictures whose fill sits inside alternate content, as Mac PowerPoint saves a PDF picture with a PNG fallback, now draw the branch the renderer supports instead of being left out.
- 8b61ea4: The PowerPoint engine can list the font faces a slide or deck asks for before layout, so hosts can load exactly those fonts before the first paint. Rendering is unchanged.
- 1398cff: Default fonts now load CJK script coverage for unmapped families and mixed text when the CJK packs are available; other scripts keep the browser font. `requestedFontFacesJson` adds optional `scripts` per face, and presentations with only coverage fonts render successfully.
- dc78510: A family with no bundled face now takes PowerPoint's substitute by name, as the agents package already does, in default font sessions. It no longer depends on which face registered first, which could also change saved autofit.
- 4f03ee2: Text drawn with a substitute face stretches each piece up to 1.45 times (was 1.33), closing more gaps in wide letters such as Century Gothic and Arial Rounded headings. Word fitting of other text is unchanged.
- 4310ba5: Text drawn with a substitute face (Aptos, Verdana, Century Gothic) no longer gaps inside words: each piece the layout places is stretched to its span, alike on canvas and native. Display-list runs shaped without kerning carry `kerning: false`, and both engines draw them unkerned.
- 47edf38: Slide text stored with CR LF or a lone CR reads as one line break, as in PowerPoint, so an edited paragraph no longer saves a stray carriage return and the text can be written back unchanged.
- 76bb148: Native slide and page PNGs draw a PNG picture that ends without its closing chunk, as browsers do, instead of skipping it.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- f13ed22: SVG and metafile pictures in native slide PNGs render within a bounded cost per picture and per slide, drawing at a reduced scale when a picture would cost more.
- fd25d97: Text in the Symbol font now draws the Greek letters, arrows and math signs Office shows for it, in documents, slides and cells alike. Wingdings and Webdings run text uses the same shared slot table as bullets. Saved text is unchanged.
- 1175683: TIFF pictures whose decoded size their dimensions don't bound, such as JPEG-compressed ones, are no longer decoded, keeping memory use predictable. Adds `tiff_bounded_pixels`. DOCX opens charge a per-document TIFF pixel budget. VSDX rendering skips TIFF pictures.
- Updated dependencies [f99e53a]
- Updated dependencies [de8f092]
- Updated dependencies [1a82d79]
- Updated dependencies [3609ce4]
- Updated dependencies [5999962]
- Updated dependencies [57f4975]
- Updated dependencies [d60b8e4]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [b8395eb]
- Updated dependencies [d60b8e4]
- Updated dependencies [53502c6]
- Updated dependencies [ec04baa]
- Updated dependencies [d60b8e4]
- Updated dependencies [21a60e5]
- Updated dependencies [19c20c3]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [18b3628]
- Updated dependencies [dc78510]
- Updated dependencies [9ac1ffe]
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts@0.5.0
  - @betteroffice/fonts-cjk@0.5.0
  - @betteroffice/fonts-intl@0.5.0
  - @betteroffice/fonts-latin@0.5.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
