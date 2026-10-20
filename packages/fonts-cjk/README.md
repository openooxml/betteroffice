# @betteroffice/fonts-cjk

Chinese, Japanese and Korean faces for [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts).
The five faces add 33 MB, so they ship as a separate, optional package.

## Install

```bash
bun add @betteroffice/fonts @betteroffice/fonts-cjk
```

## Usage

No import from this package is needed. Install it next to `@betteroffice/fonts` and configure that
package once; it loads these faces through an optional dynamic import.

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts });
```

The `@betteroffice/fonts/cdn` entry loads the CJK faces from a CDN without installing this package.

## Contents

| Bundled family | Substitutes for (Word families)                                                                                   | Script bucket | License | Version |
| -------------- | ----------------------------------------------------------------------------------------------------------------- | ------------- | ------- | ------- |
| Noto Sans SC   | Microsoft YaHei, SimHei, DengXian (微软雅黑, 黑体, 等线)                                                          | `cjk-sc`      | OFL 1.1 | 2.004   |
| Noto Serif SC  | SimSun, NSimSun, FangSong, KaiTi (宋体, 仿宋, 楷体)                                                               | `cjk-sc`      | OFL 1.1 | 2.003   |
| Noto Sans TC   | Microsoft JhengHei, PMingLiU, MingLiU, DFKai-SB (微軟正黑體, 新細明體, 細明體, 標楷體)                            | `cjk-tc`      | OFL 1.1 | 2.004   |
| Noto Sans JP   | MS (P)Gothic, MS (P)Mincho, Meiryo, Yu Gothic, Yu Mincho (ＭＳ ゴシック, ＭＳ 明朝, メイリオ, 游ゴシック, 游明朝) | `cjk-jp`      | OFL 1.1 | 2.004   |
| Noto Sans KR   | Malgun Gothic, Gulim, Dotum, Batang, Gungsuh (맑은 고딕, 굴림, 돋움, 바탕, 궁서)                                  | `cjk-kr`      | OFL 1.1 | 2.004   |

Each face ships Regular. Ming, Mincho and Batang families map to the regional Noto Sans faces. The
files are noto-cjk's static `SubsetOTF` Regulars, so the browser and the engine measure the same
advances.

## API

```ts
import { CJK_FONT_ASSET_URLS } from '@betteroffice/fonts-cjk';
```

`CJK_FONT_ASSET_URLS` maps each asset filename, as named in the `@betteroffice/fonts` manifest, to a
`() => URL` resolver. It exists so bundlers emit the assets; hosts do not normally import it.

## Documentation

- [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts)
- [Packages and fonts](https://docs.betteroffice.dev/docs/packages#fonts)

## Licensing

The loader code is Apache-2.0 (see `LICENSE`). The font binaries are licensed under the SIL Open Font License 1.1; `LICENSES/OFL-NotoCJK.txt` carries the full license text preceded by the copyright notices it applies to. Notices: Copyright 2014-2021 Adobe for Noto Sans SC/TC/JP/KR, and Copyright 2017-2024 Adobe for Noto Serif SC, both with Reserved Font Name "Source". Static `SubsetOTF` Regulars vendored from [notofonts/noto-cjk](https://github.com/notofonts/noto-cjk) at commit `f8d157532fbfaeda587e826d4cd5b21a49186f7c` (`Sans/SubsetOTF/{SC,TC,JP,KR}/`, `Serif/SubsetOTF/SC/`).
