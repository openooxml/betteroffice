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
compiled into the `betteroffice-xlsx-raster` crate via `include_bytes!`.

Copyright 2013 The Carlito Project Authors, with Reserved Font Name "Carlito".

License: SIL Open Font License, Version 1.1
(`crates/xlsx-raster/assets/OFL.txt`).

---

## Bundled font binaries (`packages/fonts`, `packages/fonts-cjk`)

Thirty font binaries are vendored unmodified and redistributed in the published
`@betteroffice/fonts` and `@betteroffice/fonts-cjk` packages. All are licensed
under the SIL Open Font License, Version 1.1. Each package ships the applicable
license texts, with the copyright notices they cover, under its `LICENSES/`
directory; those files are the authoritative copies for redistribution.

`packages/fonts/assets` — 25 faces:

- **Carlito** (4 faces) — Copyright 2013 The Carlito Project Authors, with
  Reserved Font Name "Carlito". From google/fonts `ofl/carlito`; upstream
  https://github.com/googlefonts/carlito. `LICENSES/OFL-Carlito.txt`.
- **Caladea** (4 faces) — Copyright 2012 The Caladea Project Authors. From
  google/fonts `ofl/caladea`; upstream
  https://github.com/huertatipografica/Caladea. `LICENSES/OFL-Caladea.txt`.
- **Liberation Sans / Serif / Mono** (12 faces) — Digitized data copyright (c)
  2010 Google Corporation; Copyright (c) 2012 Red Hat, Inc., with Reserved Font
  Name "Liberation". From the Liberation Fonts 2.1.5 release,
  https://github.com/liberationfonts/liberation-fonts.
  `LICENSES/OFL-Liberation.txt`.
- **Noto Sans Arabic, Noto Naskh Arabic** (3 faces) — Copyright 2022 The Noto
  Project Authors. From https://github.com/notofonts/notofonts.github.io;
  upstream https://github.com/notofonts/arabic. `LICENSES/OFL-NotoArabic.txt`.
- **Noto Sans Hebrew** (2 faces) — Copyright 2022 The Noto Project Authors.
  From https://github.com/notofonts/notofonts.github.io; upstream
  https://github.com/notofonts/hebrew. `LICENSES/OFL-NotoSansHebrew.txt`.

`packages/fonts-cjk/assets` — 5 faces:

- **Noto Sans SC / TC / JP / KR** (4 faces) — Copyright 2014-2021 Adobe, with
  Reserved Font Name "Source".
- **Noto Serif SC** (1 face) — Copyright 2017-2024 Adobe, with Reserved Font
  Name "Source".

Both are static `SubsetOTF` Regulars from
https://github.com/notofonts/noto-cjk. `LICENSES/OFL-NotoCJK.txt`.
