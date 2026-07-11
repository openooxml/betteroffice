# Third-party notices

This file records third-party software included in, or from which substantial
portions of, this repository derive, together with the applicable licenses.
Entries are appended as new third-party material is incorporated.

---

## eigenpal docx editor

The API design and portions of the TypeScript packages in this repository
derive from the eigenpal docx editor (upstream repository:
https://github.com/eigenpal/docx-editor, mirrored at
https://github.com/sorenlouv/docx-editor), published on npm as the following
packages (last release: 1.9.0):

- `@eigenpal/docx-editor-core`
- `@eigenpal/docx-editor-react`
- `@eigenpal/docx-editor-vue`
- `@eigenpal/docx-editor-i18n`
- `@eigenpal/docx-editor-agents`
- `@eigenpal/nuxt-docx-editor`

License: Apache License, Version 2.0 (per the `license` field of the published
package metadata and the `LICENSE` file shipped in the packages).
Copyright 2026 EigenPal Inc.

The license terms are identical to this repository's root `LICENSE`; that
file serves as the copy of the license for this derivation.

---

## Carlito (metric-compatible with Calibri)

`crates/xlsx-raster/assets/Carlito-Regular.ttf` is vendored unmodified from
the google/fonts repository, `ofl/carlito` (https://github.com/google/fonts);
upstream project https://github.com/googlefonts/carlito. The bytes are
compiled into the `xlsx-raster` crate via `include_bytes!`.

Copyright 2013 The Carlito Project Authors, with Reserved Font Name "Carlito".

License: SIL Open Font License, Version 1.1
(`crates/xlsx-raster/assets/OFL.txt`).
