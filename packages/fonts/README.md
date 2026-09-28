# @betteroffice/fonts

Bundled open fonts for the [BetterOffice](https://betteroffice.dev/) engine, plus a small lazy loader: metric-compatible Latin replacements for the MS core fonts, and script-coverage faces for CJK and RTL text.

This package ships open metric-compatible substitutes for the MS core fonts (Calibri, Cambria, Arial, Times New Roman, Courier New): same advance widths, so line breaks and pagination match Word.

Beyond the core replacements it also ships the families PowerPoint decks are
commonly drawn with — Inter, Roboto, Source Sans 3, DM Sans, DM Serif Display,
Open Sans, Montserrat, Poppins, Oswald and Heebo, the fonts themselves — plus
Gelasio for Georgia and Comic Relief for Comic Sans MS, both metric-compatible.
Unknown families resolve to a related, monospace, serif or sans face: Calibri, or Arial for Word.

## Why it matters

Measured across 813 real-world documents, scored against Word's own page count in `docProps/app.xml <Pages>`:

|                            | exact page-count match | within ±1 | mean abs error |
| -------------------------- | ---------------------- | --------- | -------------- |
| This package configured    | **61.9%**              | 84.6%     | 0.80           |
| No font provider at all    | 46.5%                  | 70.8%     | 2.47           |

## Install

```sh
npm install @betteroffice/fonts
```

Hand the module to `@betteroffice/docx` once, before any editor mounts:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts });
```

`configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` does the same lazily, keeping the package in its own chunk. To leave the binaries off your origin, add a `baseUrl` — see [Serving the faces from a CDN](#serving-the-faces-from-a-cdn).

Add [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk) when your documents contain Chinese, Japanese or Korean text — those five faces are 33 MB and ship separately so nobody installs them unnecessarily.

| Package                   | Faces | Size on disk | Contents                                                   |
| ------------------------- | ----- | ------------ | ---------------------------------------------------------- |
| `@betteroffice/fonts`     | 65    | 14.2 MB      | Latin metric-compatible set, deck families + Hebrew/Arabic |
| `@betteroffice/fonts-cjk` | 5     | 33 MB        | Noto Sans SC/TC/JP/KR, Noto Serif SC                       |

Faces load lazily. Regular and bold Calibri use Carlito Regular and Bold.

## Metric-compatibility mapping (Latin)

| Bundled family   | Metric-compatible with | Aliases also resolved | License | Version |
| ---------------- | ---------------------- | --------------------- | ------- | ------- |
| Carlito          | Calibri                | —                     | OFL 1.1 | 1.104   |
| Caladea          | Cambria                | —                     | OFL 1.1 | 1.001   |
| Liberation Sans  | Arial                  | Helvetica             | OFL 1.1 | 2.1.5   |
| Liberation Serif | Times New Roman        | Times                 | OFL 1.1 | 2.1.5   |
| Liberation Mono  | Courier New            | Courier               | OFL 1.1 | 2.1.5   |

Gelasio 1.008 is metric-compatible with Georgia, and Comic Relief 1.200 with Comic Sans MS; both are OFL 1.1. Comic Relief ships Regular and Bold; the other six families ship Regular, Bold, Italic and BoldItalic.

The deck families (Inter, Roboto, Source Sans 3 for Source Sans Pro, DM Sans, Open Sans, Montserrat, Poppins) ship the same four faces. Oswald and Heebo ship Regular and Bold, DM Serif Display ships Regular and Italic.

## Script-coverage mapping (CJK + RTL)

These Noto faces provide CJK and RTL glyph coverage for the Rust engine and browser.

For the bundled entry, install `@betteroffice/fonts-cjk` for the CJK faces below. RTL faces ship in this package.

| Bundled family    | Substitutes for (Word families)                                                                                   | Script bucket | License | Version |
| ----------------- | ----------------------------------------------------------------------------------------------------------------- | ------------- | ------- | ------- |
| Noto Sans SC      | Microsoft YaHei, SimHei, DengXian (微软雅黑, 黑体, 等线)                                                          | `cjk-sc`      | OFL 1.1 | 2.004   |
| Noto Serif SC     | SimSun, NSimSun, FangSong, KaiTi (宋体, 仿宋, 楷体)                                                               | `cjk-sc`      | OFL 1.1 | 2.003   |
| Noto Sans TC      | Microsoft JhengHei, PMingLiU, MingLiU, DFKai-SB (微軟正黑體, 新細明體, 細明體, 標楷體)                            | `cjk-tc`      | OFL 1.1 | 2.004   |
| Noto Sans JP      | MS (P)Gothic, MS (P)Mincho, Meiryo, Yu Gothic, Yu Mincho (ＭＳ ゴシック, ＭＳ 明朝, メイリオ, 游ゴシック, 游明朝) | `cjk-jp`      | OFL 1.1 | 2.004   |
| Noto Sans KR      | Malgun Gothic, Gulim, Dotum, Batang, Gungsuh (맑은 고딕, 굴림, 돋움, 바탕, 궁서)                                  | `cjk-kr`      | OFL 1.1 | 2.004   |
| Noto Sans Hebrew  | — (Hebrew coverage)                                                                                               | `hebrew`      | OFL 1.1 | 3.001   |
| Noto Sans Arabic  | — (Arabic coverage)                                                                                               | `arabic`      | OFL 1.1 | 2.013   |
| Noto Naskh Arabic | — (serif Arabic face)                                                                                             | `arabic`      | OFL 1.1 | 2.021   |

Notes:

- CJK faces ship Regular. Ming, Mincho and Batang families map to regional Noto Sans faces.
- The CJK binaries are noto-cjk's static `SubsetOTF` Regulars, so browser and Rust measurement read identical advances.
- The script chain supplies Hebrew and Arabic glyphs alongside Latin mappings. Hebrew and Arabic sans ship Regular and Bold.

## Why raw TTF (sfnt), not woff2

The same bytes are consumed by two sides at once:

- the **browser**, via `registerBundledFontFace()` (`FontFace` API), so DOM text measurement uses these exact bytes;
- the **Rust/WASM `FontStore`**, via `loadBundledFontBytes()`, which parses raw sfnt.

Byte-identity across both consumers is a hard requirement: the two measurement paths must be handed the same font bytes, or their results diverge.

## Lazy loading

Importing this package performs **no network activity and no font registration**. Font binaries are fetched lazily, per face, on the first `loadBundledFontBytes()` / `registerBundledFontFace()` call. The fetch is same-origin: asset URLs are derived with `new URL(..., import.meta.url)` so bundlers (Vite) emit the files alongside the module — nothing is loaded from a CDN or any remote host.

## Serving the faces from a CDN

Use the separate CDN entry to keep font binaries and the CJK add-on out of your application bundle:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';

configureDefaultFonts({ load: () => import('@betteroffice/fonts/cdn') });
```

This entry fetches faces lazily from version-pinned jsDelivr URLs that follow the installed `@betteroffice/fonts` and `@betteroffice/fonts-cjk` package versions, so the CDN revision always matches the installed code. Japanese documents can load the CJK faces without installing the CJK package. Each download has a 30-second deadline and omits cookies and referrers; failed or truncated downloads can be retried. The bundled entry retains its existing request behavior and remains available for offline applications.

To use your own CDN, pass a provider to the editor:

```ts
import { createFontProvider } from '@betteroffice/fonts/cdn';

const measurementFontProvider = createFontProvider({
  baseUrl: 'https://cdn.example.com/fonts/latin/',
  cjkBaseUrl: 'https://cdn.example.com/fonts/cjk/',
});
```

Pass `measurementFontProvider` to `DocxEditor`. When only `baseUrl` is supplied, all faces use that directory. Serve the original package filenames and allow cross-origin requests. Your content security policy must permit the CDN in `connect-src`.

The original entry also supports a custom origin, but its imports still reference package assets. It keeps same-origin loading as the default for offline and strict-CSP deployments. A CDN can observe which font assets are requested. Opt in explicitly, either through the engine:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts, baseUrl: 'https://cdn.example.com/betteroffice-fonts/' });
```

or by building the provider yourself with `createFontProvider({ baseUrl })`. The base URL is joined with each face's asset filename, so serve the contents of `assets/` at that path.

**A base URL moves the binaries, not the package.** The manifest that maps a Word font name to a face lives in this package's 23 KB of JavaScript, so a CDN deployment still installs `@betteroffice/fonts`; what it stops shipping is the 14.2 MB of faces. Pass the font module with `baseUrl`.

`configureDefaultFonts` is process-global: call it at module initialization, before any editor resolves fonts; pass per-editor providers for tenant-specific URLs.

A relative base URL is pinned to the current browser route when configuration or provider creation runs; outside a browser, the base URL must be absolute.

**A base URL bypasses package resolution entirely, including the CJK add-on.** Every face — Latin, RTL and CJK alike — is then fetched from that one base, so if your documents contain CJK you must serve `@betteroffice/fonts-cjk`'s `assets/` from the same directory. Filenames do not collide, so copying both packages' `assets/` into one folder is enough.

Every loaded asset is checked against the vendored manifest's exact decoded byte length. This detects truncation, not same-length tampering, so the configured origin still needs to be trusted.

Serve the files with `Content-Encoding: br` or `gzip`. The faces are TTF/OTF rather than woff2 (see above), and transport compression recovers most of the difference: the 65 faces are 14,527 KB raw and 7,156 KB gzipped.

## Bundler note

With rollup or esbuild, externalize the font packages to preserve their asset URLs. Alternatively, configure a font module with `baseUrl`.

```sh
esbuild app.js --bundle --packages=external
# or, narrowly:
esbuild app.js --bundle --external:@betteroffice/fonts --external:@betteroffice/fonts-cjk
```

```js
// rollup.config.js — also fixes Vite builds that inline the package
export default {
  external: ['@betteroffice/fonts', '@betteroffice/fonts-cjk'],
};
```

## Deterministic resolution

With the bundled provider available, measurement never consults OS-installed fonts. Font resolution is embedded document faces first, then the bundled metric-compatible substitutes, then the always-available last-resort base face — so the same document with the same provider measures identically on every machine.

## API

Most hosts need none of this — `configureDefaultFonts({ fonts })` is enough. It is here for custom byte sources, non-docx consumers, and browser-side `FontFace` registration.

```ts
import {
  createFontProvider, // ({ baseUrl }?) -> the provider the measurement engine consumes
  BUNDLED_FONTS, // BundledFontFace[] — the full manifest (single source of truth)
  resolveMetricCompatFamily, // "calibri" -> "Carlito" (case-insensitive, aliases included)
  resolveMetricCompatFace, // ("SimHei", bold, italic) -> concrete face (else Regular)
  resolveScriptFallbackFace, // ('cjk-sc' | 'arabic' | ..., bold, italic) -> coverage face
  resolveLastResortFace, // always-available base face for any (family, bold, italic)
  loadBundledFontBytes, // (face, { baseUrl }?) -> Promise<ArrayBuffer> (cached per face + base)
  registerBundledFontFace, // face -> FontFace registration (no-op outside the DOM)
} from '@betteroffice/fonts';
```

## Licensing

The loader code is Apache-2.0 (see `LICENSE`). The font binaries are licensed under the SIL Open Font License 1.1; the full license texts with per-family copyright notices are in `LICENSES/`:

- `LICENSES/OFL-Carlito.txt` — Copyright 2013 The Carlito Project Authors, Reserved Font Name "Carlito". Vendored from [google/fonts `ofl/carlito`](https://github.com/google/fonts/tree/main/ofl/carlito) (upstream: [googlefonts/carlito](https://github.com/googlefonts/carlito)).
- `LICENSES/OFL-Caladea.txt` — Copyright 2012 The Caladea Project Authors. Vendored from [google/fonts `ofl/caladea`](https://github.com/google/fonts/tree/main/ofl/caladea) (upstream: [huertatipografica/Caladea](https://github.com/huertatipografica/Caladea)).
- `LICENSES/OFL-Liberation.txt` — Digitized data copyright (c) 2010 Google Corporation; Copyright (c) 2012 Red Hat, Inc., Reserved Font Name Liberation. Vendored unmodified from the [Liberation Fonts 2.1.5 release](https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5).
- `LICENSES/OFL-NotoSansHebrew.txt` — Copyright 2022 The Noto Project Authors. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansHebrew/hinted/ttf/`); upstream project [notofonts/hebrew](https://github.com/notofonts/hebrew).
- `LICENSES/OFL-NotoArabic.txt` — Copyright 2022 The Noto Project Authors; covers Noto Sans Arabic and Noto Naskh Arabic. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansArabic/hinted/ttf/`, `fonts/NotoNaskhArabic/hinted/ttf/`); upstream project [notofonts/arabic](https://github.com/notofonts/arabic).
- `LICENSES/OFL-Gelasio.txt`, `OFL-ComicRelief.txt`, `OFL-Inter.txt`, `OFL-Roboto.txt`, `OFL-SourceSans3.txt`, `OFL-DMSans.txt`, `OFL-DMSerifDisplay.txt`, `OFL-OpenSans.txt`, `OFL-Montserrat.txt`, `OFL-Poppins.txt`, `OFL-Oswald.txt`, `OFL-Heebo.txt`: one per family, each with that family's copyright notice.

The CJK binaries and their `OFL-NotoCJK.txt` license text ship in [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk).
