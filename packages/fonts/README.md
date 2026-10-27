# @betteroffice/fonts

Open fonts and a lazy loader for the BetterOffice layout engines: metric-compatible substitutes for
the Microsoft core fonts, common presentation families, and Hebrew and Arabic coverage. The DOCX
engine loads them through `configureDefaultFonts`; for the PPTX editor's `fonts` prop,
`loadBundledFontBytes` returns the bytes of any bundled face. The CJK faces
ship separately in [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk).

The substitutes share the advance widths of Calibri, Cambria, Arial, Times New Roman and Courier
New, so line breaks and pagination match Word. Across 813 real-world documents, configuring this
package raised exact page-count agreement with Word from 46.5% to 61.9%.

## Install

```bash
bun add @betteroffice/fonts
# for Chinese, Japanese or Korean text:
bun add @betteroffice/fonts-cjk
```

## Usage

Configure the DOCX engine once, before any editor mounts:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts });
```

To keep the package in its own chunk, pass `load` instead:

```ts
configureDefaultFonts({ load: () => import('@betteroffice/fonts') });
```

`configureDefaultFonts` is process-global. Call it at module initialization; for per-editor or
per-tenant URLs, pass a provider to `DocxEditor` as `measurementFontProvider` instead.

Importing the package does not fetch or register anything. Each face is fetched the first time it is
needed, from the package's own assets (same origin), so bundlers such as Vite emit the files next to
the module.

## Serving the faces from a CDN

The `@betteroffice/fonts/cdn` entry keeps the font binaries and the CJK add-on out of your bundle. It
fetches faces from jsDelivr URLs pinned to the installed package versions:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';

configureDefaultFonts({ load: () => import('@betteroffice/fonts/cdn') });
```

To use your own CDN, create a provider and pass it to `DocxEditor` as `measurementFontProvider`:

```ts
import { createFontProvider } from '@betteroffice/fonts/cdn';

const measurementFontProvider = createFontProvider({
  baseUrl: 'https://cdn.example.com/fonts/latin/',
  cjkBaseUrl: 'https://cdn.example.com/fonts/cjk/',
});
```

Or keep the bundled entry and move only the binaries:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts, baseUrl: 'https://cdn.example.com/betteroffice-fonts/' });
```

- Serve the original filenames from the packages' `assets/` directories, with CORS enabled and the
  origin allowed in your `connect-src` policy.
- With a single `baseUrl`, every face, CJK included, loads from that directory. Copy the `assets/`
  of both packages into it; the filenames do not collide.
- A relative `baseUrl` resolves against the current page; outside a browser it must be absolute.
- Each download is checked against the manifest's byte length, which detects truncation but not
  tampering, so use an origin you trust. A CDN can see which faces a document requests.
- Serve the files with `br` or `gzip` encoding. The faces are raw TTF/OTF, because the browser and
  the WebAssembly engine must measure the same bytes; compression brings the 65 faces from 14,527 KB to
  7,156 KB.

## Bundlers

With rollup, esbuild or a Vite build that inlines the package, mark the font packages as external
so their asset URLs survive, or configure a `baseUrl`:

```bash
esbuild app.js --bundle --external:@betteroffice/fonts --external:@betteroffice/fonts-cjk
```

```js
// rollup.config.js
export default {
  external: ['@betteroffice/fonts', '@betteroffice/fonts-cjk'],
};
```

## Bundled faces

| Bundled family   | Metric-compatible with | Aliases also resolved | License | Version |
| ---------------- | ---------------------- | --------------------- | ------- | ------- |
| Carlito          | Calibri                | none                  | OFL 1.1 | 1.104   |
| Selawik          | Segoe UI, Light, Semilight, Semibold | none | OFL 1.1 | |
| Caladea          | Cambria                | none                  | OFL 1.1 | 1.001   |
| Liberation Sans  | Arial                  | Helvetica             | OFL 1.1 | 2.1.5   |
| Liberation Serif | Times New Roman        | Times                 | OFL 1.1 | 2.1.5   |
| Liberation Mono  | Courier New            | Courier               | OFL 1.1 | 2.1.5   |

Gelasio 1.008 is metric-compatible with Georgia and Comic Relief 1.200 with Comic Sans MS (both OFL
1.1). The presentation families Inter, Roboto, Source Sans 3 (for Source Sans Pro), DM Sans, DM Serif
Display, Open Sans, Montserrat, Poppins, Oswald and Heebo are the fonts themselves. Unknown families
resolve to a related monospace, serif or sans face. Selawik ships Regular, Bold, Light, Semilight
and Semibold upright faces. Italic Segoe UI requests keep the last-resort chain, whose italic is a
real italic face.

| Bundled family    | Coverage | Script bucket | License | Version |
| ----------------- | -------- | ------------- | ------- | ------- |
| Noto Sans Hebrew | Hebrew coverage | `hebrew` | OFL 1.1 | 3.001 |
| Noto Sans Arabic | Arabic coverage | `arabic` | OFL 1.1 | 2.013 |
| Noto Naskh Arabic | serif Arabic face | `arabic` | OFL 1.1 | 2.021 |

Measurement never consults fonts installed on the machine: embedded document fonts come first, then
the bundled substitutes, then a last-resort face, so a document measures the same everywhere.

## API

Most hosts only need `configureDefaultFonts`. The package also exports the pieces for custom byte
sources and other consumers:

| export | description |
| --- | --- |
| `createFontProvider({ baseUrl? })` | The provider the measurement engine consumes. |
| `BUNDLED_FONTS` | The face manifest. |
| `resolveMetricCompatFamily`, `resolveMetricCompatFace` | Map a Word family (and style) to its bundled substitute. |
| `resolveBundledFamilyFace` | Bundled face for a family and style, or `undefined`; missing italics use browser synthesis. |
| `resolveScriptFallbackFace` | Coverage face for a script bucket such as `cjk-sc` or `arabic`. |
| `resolveGlyphFallbackFace`, `createFontProvider().resolveGlyphFallback` | Coverage metadata and lazy bytes for characters absent from Selawik; optional `office` defaults to `powerpoint`, with `word` selecting Word's fallback. |
| `resolveLastResortFace` | The always-available fallback face. |
| `resolveDigitAdvance`, `maxDigitWidth` | Widest digit advance, recorded or measured. |
| `loadBundledFontBytes(face, { baseUrl? })` | Fetches a face's bytes, cached per face and base URL. |
| `registerBundledFontFace(face)` | Registers a face with the browser `FontFace` API; no-op outside the DOM. |

## Documentation

- [Packages and fonts](https://docs.betteroffice.dev/docs/packages#fonts)
- [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk)
- [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx)

## Licensing

The loader code is Apache-2.0 (see `LICENSE`). The font binaries are licensed under the SIL Open Font License 1.1; the full license texts with per-family copyright notices are in `LICENSES/`:

- `LICENSES/OFL-Carlito.txt`: Copyright 2013 The Carlito Project Authors, Reserved Font Name "Carlito". Vendored from [google/fonts `ofl/carlito`](https://github.com/google/fonts/tree/main/ofl/carlito) (upstream: [googlefonts/carlito](https://github.com/googlefonts/carlito)).
- `LICENSES/OFL-Selawik.txt`: Copyright 2015 Microsoft Corporation, Reserved Font Name "Selawik". Vendored unmodified from [microsoft/Selawik](https://github.com/microsoft/Selawik).
- `LICENSES/OFL-Caladea.txt`: Copyright 2012 The Caladea Project Authors. Vendored from [google/fonts `ofl/caladea`](https://github.com/google/fonts/tree/main/ofl/caladea) (upstream: [huertatipografica/Caladea](https://github.com/huertatipografica/Caladea)).
- `LICENSES/OFL-Liberation.txt`: Digitized data copyright (c) 2010 Google Corporation; Copyright (c) 2012 Red Hat, Inc., Reserved Font Name Liberation. Vendored unmodified from the [Liberation Fonts 2.1.5 release](https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5).
- `LICENSES/OFL-NotoSansHebrew.txt`: Copyright 2022 The Noto Project Authors. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansHebrew/hinted/ttf/`); upstream project [notofonts/hebrew](https://github.com/notofonts/hebrew).
- `LICENSES/OFL-NotoArabic.txt`: Copyright 2022 The Noto Project Authors; covers Noto Sans Arabic and Noto Naskh Arabic. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansArabic/hinted/ttf/`, `fonts/NotoNaskhArabic/hinted/ttf/`); upstream project [notofonts/arabic](https://github.com/notofonts/arabic).
- `LICENSES/OFL-Gelasio.txt`, `OFL-ComicRelief.txt`, `OFL-Inter.txt`, `OFL-Roboto.txt`, `OFL-SourceSans3.txt`, `OFL-DMSans.txt`, `OFL-DMSerifDisplay.txt`, `OFL-OpenSans.txt`, `OFL-Montserrat.txt`, `OFL-Poppins.txt`, `OFL-Oswald.txt`, `OFL-Heebo.txt`: one per family, each with that family's copyright notice.

The CJK binaries and their `OFL-NotoCJK.txt` license text ship in [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk).
