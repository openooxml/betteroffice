# @betteroffice/fonts

Bundled open fonts for the [BetterOffice](https://betteroffice.dev/) engine, plus a small lazy loader: metric-compatible Latin replacements for the MS core fonts, and script-coverage faces for CJK and RTL text.

Word documents overwhelmingly reference the MS core fonts (Calibri, Cambria, Arial, Times New Roman, Courier New), whose binaries cannot be redistributed. This package ships the open fonts the LibreOffice/ChromeOS ecosystem uses as drop-in metric replacements: same advance widths, so line breaks and pagination match Word even where glyph outlines differ slightly.

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

`@betteroffice/docx` never reaches for this package on its own: nothing is co-installed, no bundler has to resolve it, and measurement stays on the browser fallback until you say otherwise. Hand the module over once, before any editor mounts:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts });
```

`configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` does the same lazily, keeping the package in its own chunk. To leave the binaries off your origin, add a `baseUrl` — see [Serving the faces from a CDN](#serving-the-faces-from-a-cdn).

Add [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk) when your documents contain Chinese, Japanese or Korean text — those five faces are 33 MB and ship separately so nobody installs them unnecessarily.

| Package                   | Faces | Size on disk | Contents                                     |
| ------------------------- | ----- | ------------ | -------------------------------------------- |
| `@betteroffice/fonts`     | 25    | 7.9 MB       | Latin metric-compatible set + Hebrew/Arabic  |
| `@betteroffice/fonts-cjk` | 5     | 33 MB        | Noto Sans SC/TC/JP/KR, Noto Serif SC         |

Faces are fetched per face, lazily. A typical English document using regular and bold Calibri pulls Carlito Regular + Bold plus the chain's always-appended Liberation Sans Regular + Bold: 2,135,668 bytes (2.04 MiB) raw.

## Metric-compatibility mapping (Latin)

| Bundled family   | Metric-compatible with | Aliases also resolved | License | Version |
| ---------------- | ---------------------- | --------------------- | ------- | ------- |
| Carlito          | Calibri                | —                     | OFL 1.1 | 1.104   |
| Caladea          | Cambria                | —                     | OFL 1.1 | 1.001   |
| Liberation Sans  | Arial                  | Helvetica             | OFL 1.1 | 2.1.5   |
| Liberation Serif | Times New Roman        | Times                 | OFL 1.1 | 2.1.5   |
| Liberation Mono  | Courier New            | Courier               | OFL 1.1 | 2.1.5   |

Each Latin family ships four faces: Regular, Bold, Italic, BoldItalic — 20 TTFs under `assets/`.

## Script-coverage mapping (CJK + RTL)

These faces exist so the Rust text engine (and the browser) has real glyphs for scripts the Latin faces cannot cover. **They are coverage fallbacks first, metric approximations second** — unlike Carlito/Calibri, the Noto CJK faces do NOT share advance widths with SimSun/MS Gothic/Malgun Gothic et al. (fullwidth ideographs are uniformly 1 em everywhere, but proportional Latin runs and line heights differ), so CJK pagination approximates Word rather than matching it.

The CJK rows below are **resolvable only with `@betteroffice/fonts-cjk` installed**; their manifest entries live here (resolution policy belongs in one place) but the binaries do not. Without the add-on, a CJK request resolves to a loader that rejects, and the caller falls back — the RTL faces below ship in this package and always work.

| Bundled family    | Substitutes for (Word families)                                                                                   | Script bucket | License | Version |
| ----------------- | ----------------------------------------------------------------------------------------------------------------- | ------------- | ------- | ------- |
| Noto Sans SC      | Microsoft YaHei, SimHei, DengXian (微软雅黑, 黑体, 等线)                                                          | `cjk-sc`      | OFL 1.1 | 2.004   |
| Noto Serif SC     | SimSun, NSimSun, FangSong, KaiTi (宋体, 仿宋, 楷体)                                                               | `cjk-sc`      | OFL 1.1 | 2.003   |
| Noto Sans TC      | Microsoft JhengHei, PMingLiU, MingLiU, DFKai-SB (微軟正黑體, 新細明體, 細明體, 標楷體)                            | `cjk-tc`      | OFL 1.1 | 2.004   |
| Noto Sans JP      | MS (P)Gothic, MS (P)Mincho, Meiryo, Yu Gothic, Yu Mincho (ＭＳ ゴシック, ＭＳ 明朝, メイリオ, 游ゴシック, 游明朝) | `cjk-jp`      | OFL 1.1 | 2.004   |
| Noto Sans KR      | Malgun Gothic, Gulim, Dotum, Batang, Gungsuh (맑은 고딕, 굴림, 돋움, 바탕, 궁서)                                  | `cjk-kr`      | OFL 1.1 | 2.004   |
| Noto Sans Hebrew  | — (script fallback only)                                                                                          | `hebrew`      | OFL 1.1 | 3.001   |
| Noto Sans Arabic  | — (script fallback only)                                                                                          | `arabic`      | OFL 1.1 | 2.013   |
| Noto Naskh Arabic | — (script fallback only; serif Arabic, addressable as a family)                                                   | `arabic`      | OFL 1.1 | 2.021   |

Notes:

- **Regular only (CJK).** The CJK faces ship a single Regular each; a bold CJK request resolves to the Regular face and bold falls back through the measurement font chain. Serif TC/JP/KR are not vendored (size budget) — the Ming/Mincho/Batang serif families map to the regional sans face, diverging from `fontResolver.ts`'s Noto Serif picks for those regions; coverage wins over style.
- **Static CFF, not the variable TTFs.** The CJK binaries are the static `SubsetOTF` Regulars from noto-cjk, NOT the google/fonts variable TTFs: those VFs default to the Thin (wght=100) instance, and the Rust `FontStore` reads default-instance advances while the browser measures at wght=400 — same bytes, different numbers. The statics keep both sides identical (skrifa parses CFF; verified against the measure pipeline).
- **RTL faces carry no Word-family mapping.** Hebrew/Arabic documents mostly name Latin families (Arial, Times New Roman) which keep their Liberation mapping; the per-script fallback chain supplies the Hebrew/Arabic glyphs. Hebrew and Arabic sans ship Regular + Bold statics.

## Why raw TTF (sfnt), not woff2

The same bytes are consumed by two sides at once:

- the **browser**, via `registerBundledFontFace()` (`FontFace` API), so DOM text measurement uses these exact bytes;
- the **Rust/WASM `FontStore`**, via `loadBundledFontBytes()`, which parses raw sfnt.

Byte-identity across both consumers is a hard requirement: the two measurement paths must be handed the same font bytes, or their results diverge.

## Lazy loading

Importing this package performs **no network activity and no font registration**. Font binaries are fetched lazily, per face, on the first `loadBundledFontBytes()` / `registerBundledFontFace()` call. The fetch is same-origin: asset URLs are derived with `new URL(..., import.meta.url)` so bundlers (Vite) emit the files alongside the module — nothing is loaded from a CDN or any remote host.

## Serving the faces from a CDN

Same-origin is the default deliberately: a CDN default would leak document-font usage to a third party and break offline and strict-CSP deployments. Opt in explicitly, either through the engine:

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts, baseUrl: 'https://cdn.example.com/betteroffice-fonts/' });
```

or by building the provider yourself with `createFontProvider({ baseUrl })`. The base URL is joined with each face's asset filename, so serve the contents of `assets/` at that path.

**A base URL moves the binaries, not the package.** The manifest that maps a Word font name to a face lives in this package's 14 KB of JavaScript, so a CDN deployment still installs `@betteroffice/fonts`; what it stops shipping is the 7.9 MB of faces. A `baseUrl` with no module loads nothing and warns.

`configureDefaultFonts` is process-global: call it at module initialization before any editor resolves fonts, not from `useEffect`; existing registries retain their provider, and a multi-tenant server cannot use it to choose different base URLs per tenant.

A relative base URL is pinned to the current browser route when configuration or provider creation runs; outside a browser, the base URL must be absolute.

**A base URL bypasses package resolution entirely, including the CJK add-on.** Every face — Latin, RTL and CJK alike — is then fetched from that one base, so if your documents contain CJK you must serve `@betteroffice/fonts-cjk`'s `assets/` from the same directory. Filenames do not collide, so copying both packages' `assets/` into one folder is enough.

Every loaded asset is checked against the vendored manifest's exact decoded byte length. This detects truncation, not same-length tampering, so the configured origin still needs to be trusted.

Serve the files with `Content-Encoding: br` or `gzip`. The faces are TTF/OTF rather than woff2 (see above), and transport compression recovers most of the difference: the Latin set is 7,252 KB raw and 3,602 KB gzipped.

## Bundler note

One optional edge is left: `@betteroffice/fonts` → `@betteroffice/fonts-cjk`, a dynamic `import()` of an optional peer dependency, caught and degraded at runtime when the add-on is absent. `@betteroffice/docx` no longer names this package at all, so nothing has to resolve it to build.

Measured with the add-on **not** installed, esbuild 0.28 builds cleanly and leaves the unresolved specifier in the output for the runtime catch. That holds only while the `await import()` is the direct body of the try — put it behind a conditional and esbuild resolves it eagerly and fails with exit 1, which is exactly how this package's own edge used to break consumer builds. Vite and Turbopack are clean; webpack and rollup warn.

### rollup and esbuild must not bundle this package

Separately from the above, and **regardless of whether the CJK add-on is installed**: rollup and esbuild have no asset pipeline for `new URL('../assets/…', import.meta.url)`, which is how every face locates its binary. If they inline `@betteroffice/fonts` into your output, `import.meta.url` starts pointing at your bundle, `../assets/` misses, and **every font load fails** — while the provider itself still resolves.

That half-working state is the worst of both worlds: native measurement becomes unsupported, browser hosts fall back to `measureText` and OS fonts, and pagination can vary even though the package is installed. Vite and Turbopack handle these URLs correctly and need nothing.

**Under rollup or esbuild, marking the font packages external is required for working font bytes — it is not cosmetic.**

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

Verified against the packed tarballs — externalized, both bundlers load 628,032 bytes of Carlito and 8,331,336 bytes of Noto Sans SC; inlined, both fail every load.

`configureDefaultFonts({ fonts, baseUrl })` sidesteps the asset URLs altogether by fetching every face over HTTP, which removes the reason to externalize this package at all.

## Deterministic resolution

With the bundled provider available, measurement never consults OS-installed fonts. Font resolution is embedded document faces first, then the bundled metric-compatible substitutes, then the always-available last-resort base face — so the same document with the same provider measures identically on every machine.

## API

Most hosts need none of this — `configureDefaultFonts({ fonts })` is enough. It is here for custom byte sources, non-docx consumers, and browser-side `FontFace` registration.

```ts
import {
  createFontProvider, // ({ baseUrl }?) -> the provider the measurement engine consumes
  BUNDLED_FONTS, // BundledFontFace[] — the full manifest (single source of truth)
  resolveMetricCompatFamily, // "calibri" -> "Carlito" (case-insensitive, aliases included)
  resolveMetricCompatFace, // ("SimHei", bold, italic) -> concrete face (Regular fallback)
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

The CJK binaries and their `OFL-NotoCJK.txt` license text ship in [`@betteroffice/fonts-cjk`](https://www.npmjs.com/package/@betteroffice/fonts-cjk).
