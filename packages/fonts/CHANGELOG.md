# @betteroffice/fonts

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
