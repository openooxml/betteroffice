# @betteroffice/fonts-cjk

Optional CJK font binaries for [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts). Install it alongside the base package when your documents contain Chinese, Japanese or Korean text:

```sh
npm install @betteroffice/fonts @betteroffice/fonts-cjk
```

Nothing else is required. `@betteroffice/fonts` picks these faces up through an optional dynamic import, so the CJK script-fallback chain starts resolving as soon as the package is present.

Without it nothing breaks: a CJK face resolves to a loader that rejects, the engine logs it once and measures the run with the Latin last-resort face instead. That keeps pagination on the native path, but the ideographs are measured with a font that has no glyphs for them, so CJK line breaking will be wrong.

## Why a separate package

npm has no partial-tarball fetch: a subpath export inside `@betteroffice/fonts` would still put every byte of these faces into every consumer's `node_modules`. Only a package boundary keeps them out.

| Package                   | Faces | Size on disk |
| ------------------------- | ----- | ------------ |
| `@betteroffice/fonts`     | 25    | 7.9 MB       |
| `@betteroffice/fonts-cjk` | 5     | 33 MB        |

The overwhelming majority of documents need only the base package. Measured over 100 real-world English documents, Calibri appears in 91%, Times New Roman in 79%, Arial in 56% and Cambria in 52% — all covered by the base package's metric-compatible Latin set.

## Contents

| Bundled family | Substitutes for (Word families)                                                                                   | Script bucket | License | Version |
| -------------- | ----------------------------------------------------------------------------------------------------------------- | ------------- | ------- | ------- |
| Noto Sans SC   | Microsoft YaHei, SimHei, DengXian (微软雅黑, 黑体, 等线)                                                          | `cjk-sc`      | OFL 1.1 | 2.004   |
| Noto Serif SC  | SimSun, NSimSun, FangSong, KaiTi (宋体, 仿宋, 楷体)                                                               | `cjk-sc`      | OFL 1.1 | 2.003   |
| Noto Sans TC   | Microsoft JhengHei, PMingLiU, MingLiU, DFKai-SB (微軟正黑體, 新細明體, 細明體, 標楷體)                            | `cjk-tc`      | OFL 1.1 | 2.004   |
| Noto Sans JP   | MS (P)Gothic, MS (P)Mincho, Meiryo, Yu Gothic, Yu Mincho (ＭＳ ゴシック, ＭＳ 明朝, メイリオ, 游ゴシック, 游明朝) | `cjk-jp`      | OFL 1.1 | 2.004   |
| Noto Sans KR   | Malgun Gothic, Gulim, Dotum, Batang, Gungsuh (맑은 고딕, 굴림, 돋움, 바탕, 궁서)                                  | `cjk-kr`      | OFL 1.1 | 2.004   |

Notes:

- **Coverage fallbacks first, metric approximations second.** Unlike Carlito/Calibri, the Noto CJK faces do NOT share advance widths with SimSun/MS Gothic/Malgun Gothic et al. (fullwidth ideographs are uniformly 1 em everywhere, but proportional Latin runs and line heights differ), so CJK pagination approximates Word rather than matching it.
- **Regular only.** Each face ships a single Regular; a bold CJK request resolves to the Regular face and bold falls back through the measurement font chain. Serif TC/JP/KR are not vendored (size budget) — the Ming/Mincho/Batang serif families map to the regional sans face; coverage wins over style.
- **Static CFF, not the variable TTFs.** These are the static `SubsetOTF` Regulars from noto-cjk, NOT the google/fonts variable TTFs: those VFs default to the Thin (wght=100) instance, and the Rust `FontStore` reads default-instance advances while the browser measures at wght=400 — same bytes, different numbers. The statics keep both sides identical.

## API

You do not normally import this package. `@betteroffice/fonts` resolves it on its own; the single export exists so bundlers emit the assets.

```ts
import { CJK_FONT_ASSET_URLS } from '@betteroffice/fonts-cjk';
```

`CJK_FONT_ASSET_URLS` maps an asset filename to a `() => URL` resolver, keyed exactly as `@betteroffice/fonts`' face manifest names the files.

## Licensing

The loader code is Apache-2.0 (see `LICENSE`). The font binaries are licensed under the SIL Open Font License 1.1; `LICENSES/OFL-NotoCJK.txt` carries the full licence text preceded by the copyright notices it applies to — Copyright 2014-2021 Adobe for Noto Sans SC/TC/JP/KR, and Copyright 2017-2024 Adobe for Noto Serif SC, both with Reserved Font Name "Source". Static `SubsetOTF` Regulars vendored from [notofonts/noto-cjk](https://github.com/notofonts/noto-cjk) at commit `f8d157532fbfaeda587e826d4cd5b21a49186f7c` (`Sans/SubsetOTF/{SC,TC,JP,KR}/`, `Serif/SubsetOTF/SC/`).
