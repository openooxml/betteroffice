# @betteroffice/fonts

## 0.3.0

### Minor Changes

- 2e47c0a: fonts: ship Inter, Roboto, Source Sans 3, DM Sans, Open Sans, Montserrat, Poppins, Oswald, Heebo and DM Serif Display, plus Gelasio for Georgia and Comic Relief for Comic Sans MS, so decks written in them draw in their own faces.

## 0.2.0

### Minor Changes

- 1d830df: Add a CDN-only font provider, settle Japanese font preflight without retry loops, and preserve floating header shapes without inflating body margins. Load and save alternate main-document filenames through their package relationships, and forward layout failures through the editor error callback.

### Patch Changes

- 2c658b6: Improve DOCX pagination, list formatting, justified text, header and footer spacing, anchored shapes, content-control text, and table geometry to better match Word. Use Carlito as the related fallback for Calibri Light.
- 1f5892e: Pin the CDN entry's jsDelivr URLs to the installed package versions instead of a hardcoded release.

## 0.1.0

### Minor Changes

- 6be0c18: Bundled metric-compatible fonts ship as `@betteroffice/fonts`, plus `@betteroffice/fonts-cjk` for Chinese, Japanese or Korean, and DOCX uses them only when you hand the module over: `configureDefaultFonts({ fonts })`, or `configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` to keep it in its own chunk. Installing the packages alone does nothing — without that call the engine reaches for no font package, measurement falls back to the browser, and pagination will not match Word. Because `@betteroffice/docx` no longer names `@betteroffice/fonts` anywhere in its published bundle, an esbuild consumer without the optional peer builds again.
