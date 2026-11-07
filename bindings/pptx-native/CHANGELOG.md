# @betteroffice/pptx-native

## 0.6.0

### Minor Changes

- b2ea7f5: The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.
- a45ba76: `renderPng({ slide }, options?)` on the presentation handle renders through the native rasterizer, with the same targets, options and errors as the browser.

### Patch Changes

- fa2717d: Every engine now reads decimal numbers back exactly. Drawing groups whose scaled child coordinates have long decimals stay intact on save.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- 3418edd: `renderPng` in the browser lays slides out with the bundled faces the native engine uses, so line breaks and row heights match native. On both engines, a needed face that cannot load now rejects with `PngExportError` (`font-unavailable`) naming the face.
- cade718: Kanji in Japanese decks now load and paint with Noto Sans JP instead of also fetching Noto Sans SC: Han text takes its CJK face from the run, paragraph mark or style language (Korean, Traditional and Simplified Chinese likewise).
- 0530412: The same font bytes registered under several family names are now held once and no longer count against the 256-face limit. Opening a deck with more than 256 distinct faces now fails with `PptxResourceLimitError` instead of a closed worker connection.
- e664196: The PowerPoint engine is about 3 KB smaller compressed: the slide layout cache compares slide floats with a direct walk instead of a generic serializer. Layout and cache behaviour are unchanged.
- 7f42859: `renderPng` works with no fonts passed: it uses the deck's embedded faces and the installed font packs, and honours `options.fonts`. `layoutSlide` results are unchanged by an export.
- ec04baa: `@betteroffice/pptx-native` now installs `@betteroffice/fonts-latin`, `@betteroffice/fonts-intl` and `@betteroffice/fonts-cjk` and reads their faces from disk on demand, so Latin, international and CJK text renders with real faces without any fetch. On by default.
- db4ab57: Now exposes exactly the `@betteroffice/pptx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openPresentation(Buffer)`/`Presentation` facade is removed.
- 8b61ea4: The PowerPoint engine can list the font faces a slide or deck asks for before layout, so hosts can load exactly those fonts before the first paint. Rendering is unchanged.
- 1398cff: Default fonts now load CJK script coverage for unmapped families and mixed text when the CJK packs are available; other scripts keep the browser font. `requestedFontFacesJson` adds optional `scripts` per face, and presentations with only coverage fonts render successfully.
- dc78510: A family with no bundled face now takes PowerPoint's substitute by name, as the agents package already does, in default font sessions. It no longer depends on which face registered first, which could also change saved autofit.
- 47edf38: Slide text stored with CR LF or a lone CR reads as one line break, as in PowerPoint, so an edited paragraph no longer saves a stray carriage return and the text can be written back unchanged.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- f13ed22: SVG and metafile pictures in native slide PNGs render within a bounded cost per picture and per slide, drawing at a reduced scale when a picture would cost more.
- Updated dependencies [f99e53a]
- Updated dependencies [de8f092]
- Updated dependencies [3609ce4]
- Updated dependencies [5999962]
- Updated dependencies [d60b8e4]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
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
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts@1.0.0
  - @betteroffice/fonts-cjk@1.0.0
  - @betteroffice/fonts-intl@1.0.0
  - @betteroffice/fonts-latin@1.0.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
