# @betteroffice/docx-i18n

UI locale strings, types, and runtime helpers for the
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react)
editor. `en` is the source of truth; community locales mirror its shape and fall
back to English for any untranslated key.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/docx-i18n
```

## Usage

Pass a typed locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/docx-i18n';

<DocxEditor documentBuffer={file} i18n={de} />;
```

Override individual strings by spreading a locale:

```ts
import { de } from '@betteroffice/docx-i18n';

const myLocale = { ...de, toolbar: { ...de.toolbar, bold: 'Fettdruck' } };
```

Keys set to `null` in any locale fall back to English.

## Locales

`en` (source), `de`, `he`, `pl`, `pt-BR`, `tr`, `zh-CN`. BCP-47 tags use
camelCase identifiers (`ptBR`, `zhCN`).

Each locale also ships as its own subpath (`@betteroffice/docx-i18n/pl`) so an
app that picks the locale at runtime can code-split rather than bundle them all.
For lookup by tag, `import { locales } from '@betteroffice/docx-i18n'` (pulls
every locale into the bundle).

## Types

```ts
import type {
  LocaleStrings, // shape of `en`, the full source of truth
  PartialLocaleStrings, // a community partial (null falls back)
  TranslationKey, // 'toolbar.bold' | 'dialogs.findReplace.title' | ...
  LocaleCode, // 'en' | 'de' | 'pt-BR' | ...
  TFunction,
} from '@betteroffice/docx-i18n';
```

## Non-React hosts

Build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/docx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('toolbar.bold'); // 'Fett'
```

Add keys to `en.json`, then `bun run i18n:fix` from the repo root to sync the
community locales (new keys land as `null`).

Docs: https://betteroffice.dev · Apache-2.0.
