# @betteroffice/fonts-cjk

Chinese, Japanese and Korean faces for [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts).
The optional package contains 34.82 MB of decoded fonts and 26.66 MB of packed assets.

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

`CJK_FONT_ASSET_URLS` maps each `<Face>.otf.zst` filename, as named in the
`@betteroffice/fonts` manifest, to a `() => URL` resolver with a literal asset path
for bundlers. Hosts do not normally import it.

## Assets

The package ships font pack format 1: one Zstandard frame per face and a schema 1
`assets/manifest.json` recording `codec`, `dict`, `sha256`, `encodedLength` and
decoded `byteLength`. These CFF faces use `zstd`; each family has one face and no
dictionary. The font manager decodes and verifies the assets before returning
independently owned sfnt buffers.

For a custom `baseUrl`, copy the packed assets alongside the other font packs and
serve `node_modules/@betteroffice/fonts/dist/decoder/fontpack.wasm` at
`<baseUrl>/decoder/fontpack.wasm` as `application/wasm`.

## Documentation

- [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts)
- [`@betteroffice/fonts-intl`](https://www.npmjs.com/package/@betteroffice/fonts-intl)
- [Packages and fonts](https://docs.betteroffice.dev/docs/packages#fonts)

## Licensing

Loader code is Apache-2.0 (see `LICENSE`). Fonts use the SIL Open Font License
1.1; `LICENSES/OFL-NotoCJK.txt` carries the license and copyright notices.
Static `SubsetOTF` Regulars come from
[notofonts/noto-cjk](https://github.com/notofonts/noto-cjk), with Sans pinned to
`523d033d6cb47f4a80c58a35753646f5c3608a78` and Serif to
`9b0f1436e455d902de067a2501422e5dc71ad16b`.
