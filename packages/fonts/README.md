# @betteroffice/docx-fonts

Bundled open fonts for the [@betteroffice/docx](https://betteroffice.dev/) editor, plus a small lazy loader: metric-compatible Latin replacements for the MS core fonts, and script-coverage faces for CJK and RTL text.

Word documents overwhelmingly reference the MS core fonts (Calibri, Cambria, Arial, Times New Roman, Courier New), whose binaries cannot be redistributed. This package ships the open fonts the LibreOffice/ChromeOS ecosystem uses as drop-in metric replacements: same advance widths, so line breaks and pagination match Word even where glyph outlines differ slightly.

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

Byte-identity across both consumers is a hard requirement of the differential measurement harness (see `openspec/changes/rust-canvas-engine/design.md`, "The one strategic constraint: font bytes").

## Lazy loading

Importing this package performs **no network activity and no font registration**. Font binaries are fetched lazily, per face, on the first `loadBundledFontBytes()` / `registerBundledFontFace()` call. The fetch is same-origin: asset URLs are derived with `new URL(..., import.meta.url)` so bundlers (Vite) emit the files alongside the module — nothing is loaded from a CDN or any remote host.

## Deterministic resolution

Measurement never consults OS-installed fonts. Font resolution is embedded document faces first, then the bundled metric-compatible substitutes, then the always-available last-resort base face — so the same document with the same provider measures identically on every machine.

## API

```ts
import {
  BUNDLED_FONTS, // BundledFontFace[] — the full manifest (single source of truth)
  resolveMetricCompatFamily, // "calibri" -> "Carlito" (case-insensitive, aliases included)
  resolveMetricCompatFace, // ("SimHei", bold, italic) -> concrete face (Regular fallback)
  resolveScriptFallbackFace, // ('cjk-sc' | 'arabic' | ..., bold, italic) -> coverage face
  resolveLastResortFace, // always-available base face for any (family, bold, italic)
  loadBundledFontBytes, // face -> Promise<ArrayBuffer> (cached per face)
  registerBundledFontFace, // face -> FontFace registration (no-op outside the DOM)
} from '@betteroffice/docx-fonts';
```

## Licensing

The loader code is Apache-2.0 (see `LICENSE`). The font binaries are licensed under the SIL Open Font License 1.1; the full license texts with per-family copyright notices are in `LICENSES/`:

- `LICENSES/OFL-Carlito.txt` — Copyright 2013 The Carlito Project Authors, Reserved Font Name "Carlito". Vendored from [google/fonts `ofl/carlito`](https://github.com/google/fonts/tree/main/ofl/carlito) (upstream: [googlefonts/carlito](https://github.com/googlefonts/carlito)).
- `LICENSES/OFL-Caladea.txt` — Copyright 2012 The Caladea Project Authors. Vendored from [google/fonts `ofl/caladea`](https://github.com/google/fonts/tree/main/ofl/caladea) (upstream: [huertatipografica/Caladea](https://github.com/huertatipografica/Caladea)).
- `LICENSES/OFL-Liberation.txt` — Digitized data copyright (c) 2010 Google Corporation; Copyright (c) 2012 Red Hat, Inc., Reserved Font Name Liberation. Vendored unmodified from the [Liberation Fonts 2.1.5 release](https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5).
- `LICENSES/OFL-NotoSansHebrew.txt` — Copyright 2024 The Noto Project Authors. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansHebrew/hinted/ttf/`); upstream project [notofonts/hebrew](https://github.com/notofonts/hebrew).
- `LICENSES/OFL-NotoArabic.txt` — Copyright 2022 The Noto Project Authors; covers Noto Sans Arabic and Noto Naskh Arabic. Hinted statics vendored from [notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io) at commit `cd06befda260d2abb6e5db96cf5530f80ea5180d` (`fonts/NotoSansArabic/hinted/ttf/`, `fonts/NotoNaskhArabic/hinted/ttf/`); upstream project [notofonts/arabic](https://github.com/notofonts/arabic).
- `LICENSES/OFL-NotoCJK.txt` — © 2014-2021 Adobe (Noto Sans CJK), © 2017-2024 Adobe (Noto Serif SC). Static `SubsetOTF` Regulars vendored from [notofonts/noto-cjk](https://github.com/notofonts/noto-cjk) at commit `f8d157532fbfaeda587e826d4cd5b21a49186f7c` (`Sans/SubsetOTF/{SC,TC,JP,KR}/`, `Serif/SubsetOTF/SC/`).
