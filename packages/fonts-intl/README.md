# @betteroffice/fonts-intl

Optional international Noto coverage faces for [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts).

## Install

```bash
bun add @betteroffice/fonts @betteroffice/fonts-intl
```

Configure `@betteroffice/fonts` as usual. It loads this add-on through a dynamic import when an
international face is needed.

```ts
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import * as fonts from '@betteroffice/fonts';

configureDefaultFonts({ fonts });
```

The `@betteroffice/fonts/cdn` entry can fetch the faces without installing this add-on.

## Contents

| Family | Script bucket | Styles |
| --- | --- | --- |
| Noto Sans Thai | `thai` | Regular, Bold |
| Noto Sans Lao | `lao` | Regular, Bold |
| Noto Sans Khmer | `khmer` | Regular, Bold |
| Noto Sans Myanmar | `myanmar` | Regular, Bold |
| Noto Sans Devanagari | `devanagari` | Regular, Bold |
| Noto Sans Bengali | `bengali` | Regular, Bold |
| Noto Sans Gurmukhi | `gurmukhi` | Regular, Bold |
| Noto Sans Gujarati | `gujarati` | Regular, Bold |
| Noto Sans Oriya | `oriya` | Regular, Bold |
| Noto Sans Tamil | `tamil` | Regular, Bold |
| Noto Sans Telugu | `telugu` | Regular, Bold |
| Noto Sans Kannada | `kannada` | Regular, Bold |
| Noto Sans Malayalam | `malayalam` | Regular, Bold |
| Noto Sans Sinhala | `sinhala` | Regular, Bold |
| Noto Sans Armenian | `armenian` | Regular, Bold |
| Noto Sans Georgian | `georgian` | Regular, Bold |
| Noto Sans Ethiopic | `ethiopic` | Regular, Bold |
| Noto Serif Tibetan | `tibetan` | Regular, Bold |
| Noto Sans Mongolian | `mongolian` | Regular |
| Noto Sans Syriac | `syriac` | Regular |
| Noto Sans Thaana | `thaana` | Regular, Bold |
| Noto Sans Cherokee | `cherokee` | Regular, Bold |
| Noto Sans Lao Looped | `lao` | Regular, Bold |
| Noto Sans Tifinagh | `tifinagh` | Regular |

Lao and Lao Looped share the `lao` script bucket. Script fallback prefers Noto Sans Lao;
Noto Sans Lao Looped is available by family name. Office families retain their existing family
resolution and use per-character script fallback for international glyphs.

## Assets

`INTL_FONT_ASSET_URLS` maps each `<Family>-<Style>.ttf.zst` filename to a `() => URL` resolver.
Each resolver uses a literal asset path so bundlers can emit the corresponding file.

The package ships font pack format 1: one Zstandard frame per face, plus `assets/manifest.json`.
The manifest records `codec`, `dict`, `sha256`, `encodedLength` and decoded `byteLength`.
Bold faces use their family's decoded Regular as a raw-content dictionary. The font manager
decodes the assets before supplying font bytes and returns independently owned buffers.
When using a custom `baseUrl`, copy these assets alongside the other font packages' assets.
Copy `node_modules/@betteroffice/fonts/dist/decoder/fontpack.wasm` to
`<baseUrl>/decoder/fontpack.wasm` and serve it as `application/wasm`. This decoder ships
in `@betteroffice/fonts`, and both its bundled and CDN providers use the custom root.

## Licensing

Loader code is Apache-2.0, with its license in `LICENSE`. Font binaries are SIL Open Font License
1.1, with the full text in `LICENSES/OFL-Noto.txt`. Font-specific copyright notices remain in
the decoded fonts' name tables. Hinted static TTFs come from
[notofonts/notofonts.github.io](https://github.com/notofonts/notofonts.github.io).
