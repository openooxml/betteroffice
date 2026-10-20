# @betteroffice/xlsx-i18n

UI strings, locale types and translation helpers for the
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react) editor. `en` defines
the strings; the other locales override them.

## Install

```bash
bun add @betteroffice/xlsx-i18n
```

## Usage

Pass a locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/xlsx-i18n';
import { XlsxEditor } from '@betteroffice/xlsx-react';

export function Editor({ file }: { file: Uint8Array }) {
  return <XlsxEditor file={file} i18n={de} />;
}
```

Override individual strings by spreading a locale. Keys that are missing or `null` fall back to
English.

```ts
const custom = { ...de, toolbar: { ...de.toolbar, save: 'Sichern' } };
```

## Locales

`en` (source), `de`, `fr`, `he`, `hi`, `id`, `pl`, `pt-BR`, `tr` and `zh-CN`. The
exports for `pt-BR` and `zh-CN` are `ptBR` and `zhCN`.

Each locale is also its own subpath, for example `@betteroffice/xlsx-i18n/pl`, so an app that picks
the locale at runtime can load only the one it needs. `locales` maps every tag to its strings and
bundles all of them.

## Helpers

The package exports the types `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode` and `TFunction`, and the functions `createT` and `deepMerge`.
Outside React, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/xlsx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('toolbar.save'); // 'Save'
```

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react)
- [Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/xlsx-i18n/CHANGELOG.md)
