# @betteroffice/docx-i18n

UI strings, locale types and translation helpers for the
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react) editor. `en` defines
the strings; the other locales override them.

## Install

```bash
bun add @betteroffice/docx-i18n
```

## Usage

Pass a locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/docx-i18n';
import { DocxEditor } from '@betteroffice/docx-react';

export function Editor({ file }: { file: File }) {
  return <DocxEditor documentBuffer={file} i18n={de} />;
}
```

Override individual strings by spreading a locale. Keys that are missing or `null` fall back to
English.

```ts
const custom = { ...de, formattingBar: { ...de.formattingBar, bold: 'Fettdruck' } };
```

## Locales

`en` (source), `de`, `fr`, `he`, `hi`, `id`, `pl`, `pt-BR`, `tr` and `zh-CN`. The
exports for `pt-BR` and `zh-CN` are `ptBR` and `zhCN`.

Each locale is also its own subpath, for example `@betteroffice/docx-i18n/pl`, so an app that picks
the locale at runtime can load only the one it needs. `locales` maps every tag to its strings and
bundles all of them.

## Helpers

The package exports the types `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode` and `TFunction`, and the functions `createT` and `deepMerge`.
Outside React, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/docx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('formattingBar.bold'); // 'Fett'
```

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react)
- [Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/docx-i18n/CHANGELOG.md)
