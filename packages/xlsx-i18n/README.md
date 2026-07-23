# @betteroffice/xlsx-i18n

UI locale strings, types, and runtime helpers for the
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react)
editor. `en` is the source of truth; community locales mirror its shape and fall
back to English for any untranslated key.

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

```bash
bun add @betteroffice/xlsx-i18n
```

## Usage

Pass a typed locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/xlsx-i18n';
import { XlsxEditor } from '@betteroffice/xlsx-react';

<XlsxEditor file={file} i18n={de} />;
```

Keys set to `null` in any locale fall back to English.

## Locales

`en` (source), `de`, `fr`, `he`, `hi`, `id`, `pl`, `pt-BR`, `tr`, `zh-CN`.
BCP-47 tags use camelCase identifiers (`ptBR`, `zhCN`).

Each locale also ships as its own subpath (`@betteroffice/xlsx-i18n/pl`) so an
app that picks the locale at runtime can code-split rather than bundle them all.
For lookup by tag, `import { locales } from '@betteroffice/xlsx-i18n'` (pulls
every locale into the bundle).

## Types and helpers

The package exports `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode`, `TFunction`, `createT`, and `deepMerge`. For
non-React hosts, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/xlsx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('toolbar.save');
```

Docs: https://betteroffice.dev · Apache-2.0.
