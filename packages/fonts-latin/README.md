# @betteroffice/fonts-latin

Latin and Noto RTL faces used by `@betteroffice/fonts`. Installed automatically
with the font manager. Faces load on demand.

`LATIN_FONT_ASSET_URLS` maps each `<Family>-<Style>.ttf.zst` filename to a
`() => URL` resolver with a literal asset path for bundlers. These URLs contain
font pack format 1 frames. Use `@betteroffice/fonts` to obtain decoded sfnt bytes.
`assets/manifest.json` records codecs, dictionary prerequisites, lengths and hashes.

For a custom `baseUrl`, copy the pack assets to that root and copy
`node_modules/@betteroffice/fonts/dist/decoder/fontpack.wasm` to
`<baseUrl>/decoder/fontpack.wasm`.

Loader code is Apache-2.0, with the full text in `LICENSE`. Fonts retain their
original license texts in `LICENSES`, copied from `@betteroffice/fonts`.
