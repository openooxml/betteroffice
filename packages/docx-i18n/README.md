# @betteroffice/docx-i18n

UI locale strings, types, and runtime helpers for the
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react)
editor. `en` is the source of truth; community locales mirror its shape and fall
back to English for any untranslated key.

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

```bash
bun add @betteroffice/docx-i18n
```

## Usage

Pass a typed locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/docx-i18n';
import { DocxEditor } from '@betteroffice/docx-react';

<DocxEditor documentBuffer={file} i18n={de} />;
```

Override individual strings by spreading a locale:

```ts
const myLocale = { ...de, formattingBar: { ...de.formattingBar, bold: 'Fettdruck' } };
```

Keys set to `null` in any locale fall back to English.

## Locales

`en` (source), `de`, `fr`, `he`, `hi`, `id`, `pl`, `pt-BR`, `tr`, `zh-CN`.
BCP-47 tags use camelCase identifiers (`ptBR`, `zhCN`).

Each locale also ships as its own subpath (`@betteroffice/docx-i18n/pl`) so an
app that picks the locale at runtime can code-split rather than bundle them all.
For lookup by tag, `import { locales } from '@betteroffice/docx-i18n'` (pulls
every locale into the bundle).

## Types and helpers

The package exports `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode`, `TFunction`, `createT`, and `deepMerge`. For
non-React hosts, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/docx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('formattingBar.bold'); // 'Fett'
```

Docs: https://betteroffice.dev · Apache-2.0.
