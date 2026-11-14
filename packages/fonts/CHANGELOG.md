# @betteroffice/fonts

## 0.5.0

### Minor Changes

- de8f092: The optional CJK add-on now ships compressed `.otf.zst` assets, decoded automatically by the font manager. `CJK_FONT_ASSET_URLS` keys and `BundledFontFace.file` use the packed filenames; custom asset hosts must serve them.
- 1a82d79: Without `fontProvider` or configured fonts, PNG export uses the complete bundled fonts on both engines, native included: an installed Latin or international face that cannot load rejects with `font-unavailable`. Aborting while fonts load rejects at once. `fontProvider` keeps the old behaviour. Fonts adds `isMissingFontAddon`.
- d60b8e4: The worker editor opens bundled CJK faces from their skeleton and core and fetches outline chunks for the glyphs each page draws before painting it, so a first CJK frame downloads a fraction of the face; print waits for pending outlines.
- d60b8e4: Loaders of the sliced CJK faces gain an optional `slices` location, and the new `@betteroffice/fonts/progressive` entry opens it from the skeleton and core, loading outline chunks by glyph id; calling the loader still returns the whole face, and custom providers are unaffected.
- df33f6a: Adds per-character international script fallbacks through the optional `@betteroffice/fonts-intl` package. Coverage loads on demand when the add-on is installed, preserving existing family resolution for Latin text.
- 37199d9: Ships the same Latin and RTL faces in a smaller compressed install by default. Existing providers decode faces on demand, and DOCX reuses each face across fallback registrations.
- e65df1c: Editors given `fonts` now load bundled faces for the other families a deck uses, including light members such as Calibri Light, so their text keeps its metrics. Supplied families keep the app's faces. Pass `defaultFonts={false}` to fetch no default fonts.
- 9ac1ffe: Adds `configureFontCdn` to load missing font packs from a configured CDN. Installed packs take priority, and CDN loading remains off by default.

### Patch Changes

- f99e53a: DOCX text that no loaded font covers falls back to bundled BetterOffice Base Latin faces, loaded on demand from `@betteroffice/fonts`. Native raster builds embed the faces. Private-use bullet aliases apply only to Symbol and Wingdings families.
- 3609ce4: Opening a document no longer copies a whole bundled font just to check which characters it covers, which lowers peak memory while opening CJK documents.
- 5999962: MS Mincho, MS Gothic, their proportional variants, and MS UI Gothic now use Office glyph widths by default, including Japanese names and decomposed voiced kana. CJK substitute outlines keep their original size during canvas painting.
- 57f4975: Arabic text set in Simplified Arabic or Traditional Arabic now falls back to the bundled Naskh face, closer to those fonts' style and width.
- 9ab2b61: Fonts named "Sans Serif", such as Microsoft Sans Serif, now fall back to a sans face (Arial widths in Word documents) instead of Times New Roman, so text using them measures and wraps like Word.
- 39ae19a: Packed faces fall back to the decoder shipped with the package when `<baseUrl>/decoder/fontpack.wasm` is missing, so a self-hosted `baseUrl` needs only the font assets. When no decoder loads, the error names every location tried.
- b8395eb: `checkFontPacks()` and `BundledFonts::pack_diagnostics()` report installed font packs that cannot be used, with the pack and the reason, instead of skipping them silently.
- 53502c6: Large bundled font faces decode outside the PPTX worker's engine, and registering a font in WebAssembly copies its bytes once, which lowers worker memory for CJK presentations.
- ec04baa: Reading bundled fonts from installed packages on Node.js versions without `process.getBuiltinModule` (before 20.16, or 21 to 22.2) now fails with a clear error and a one-time warning instead of a silent failed `fetch` of a `file:` URL.
- 21a60e5: The DOCX editor keeps one copy of each bundled font instead of several, which lowers memory while opening and editing documents, most for CJK documents.
- 19c20c3: The spreadsheet editor no longer keeps decoded copies of bundled fonts after the browser has registered them, which cuts page memory by about 2.5 to 5 MB.
- 842d624: Office fonts we draw with a substitute, such as Aptos, Arial Narrow, Cambria and Verdana, now measure at the real font's glyph widths and line heights, so text breaks like Word and PowerPoint. Legacy names such as Tms Rmn and Helv also resolve.
- 18b3628: New `experimentalSlicedFonts` editor option, off by default: in Chromium and Firefox it loads only the Chinese, Japanese or Korean outlines a slide needs. In `@betteroffice/fonts`, `ensureText` and `prefetch` also load the characters shaping composes text into.
- dc78510: A family with no bundled face now takes PowerPoint's substitute by name, as the agents package already does, in default font sessions. It no longer depends on which face registered first, which could also change saved autofit.
- ad45f86: Editors decode packed fonts with their engine by default instead of downloading a separate decoder. Standalone font providers can opt into a decoder with `runtime.decoder`.
- e6d5c52: Package exports now carry a `default` condition next to `import`, so webpack consumers (`next build --webpack`, including editors loaded with `next/dynamic`) resolve every entry instead of failing with "Package path is not exported".
- Updated dependencies [37199d9]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts-latin@0.5.0

## 0.4.2

### Patch Changes

- b2fa7b8: Uses Selawik for upright Segoe UI families by default, preserving their weights; italic Segoe UI keeps a real italic fallback face. Adds Office-aware `resolveGlyphFallback` and `resolveGlyphFallbackFace` coverage for characters absent from Selawik in Word and PowerPoint previews.

## 0.4.1

### Patch Changes

- c688743: CJK documents get usable fallback fonts without the CJK add-on, and `maxDigitWidth` and `resolveDigitAdvance` take optional `bold` and `italic`.

## 0.4.0

### Minor Changes

- 726c5b4: `resolveBundledFamilyFace` and `resolveFamily` also resolve a family by its bundled name (e.g. `Gelasio`), returning only faces of the requested weight and style. **BREAKING:** custom `BundledFontSource` implementations must add `resolveFamily`; the bundled providers already do.

## 0.3.0

### Minor Changes

- 2e47c0a: fonts: ship Inter, Roboto, Source Sans 3, DM Sans, Open Sans, Montserrat, Poppins, Oswald, Heebo and DM Serif Display, plus Gelasio for Georgia and Comic Relief for Comic Sans MS, so decks written in them draw in their own faces.
- 2168e58: `resolveLastResortFace` and the provider's `resolveLastResort` take an optional `office`: `'powerpoint'`, the default, substitutes Calibri for an unknown sans family and `'word'` substitutes Arial. DOCX layout follows Word.

### Patch Changes

- ed2b7be: fonts: the bundled entry now loads every face the package ships, so Inter, Roboto, Gelasio, Comic Relief and the other newly added families no longer reject with `Unknown bundled font asset`, and bundlers emit their files. A missing non-CJK face is reported as an unknown asset instead of asking for `@betteroffice/fonts-cjk`.
- 6c30f4a: Export `WORD_FAMILY_ALIASES`, the map from Word family names and their normalized spellings to the canonical Word family they stand in for (for example `helvetica` to `arial`); `resolveMetricCompatFamily` then picks the bundled face for that family.

## 0.2.0

### Minor Changes

- 1d830df: Add a CDN-only font provider, settle Japanese font preflight without retry loops, and preserve floating header shapes without inflating body margins. Load and save alternate main-document filenames through their package relationships, and forward layout failures through the editor error callback.

### Patch Changes

- 2c658b6: Improve DOCX pagination, list formatting, justified text, header and footer spacing, anchored shapes, content-control text, and table geometry to better match Word. Use Carlito as the related fallback for Calibri Light.
- 1f5892e: Pin the CDN entry's jsDelivr URLs to the installed package versions instead of a hardcoded release.

## 0.1.0

### Minor Changes

- 6be0c18: Bundled metric-compatible fonts ship as `@betteroffice/fonts`, plus `@betteroffice/fonts-cjk` for Chinese, Japanese or Korean, and DOCX uses them only when you hand the module over: `configureDefaultFonts({ fonts })`, or `configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` to keep it in its own chunk. Installing the packages alone does nothing — without that call the engine reaches for no font package, measurement falls back to the browser, and pagination will not match Word. Because `@betteroffice/docx` no longer names `@betteroffice/fonts` anywhere in its published bundle, an esbuild consumer without the optional peer builds again.
