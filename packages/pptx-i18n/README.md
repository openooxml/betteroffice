# @betteroffice/pptx-i18n

UI locale strings, types, and runtime helpers for the
[`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react)
editor. `en` defines the locale shape; community locales override its strings.

```bash
bun add @betteroffice/pptx-i18n
```

## Usage

Pass a typed locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/pptx-i18n';
import { PptxEditor } from '@betteroffice/pptx-react';

<PptxEditor file={file} fonts={fonts} i18n={de} />;
```

Keys set to `null` keep the English string.

## Locales

`en` (source), `de`, `fr`, `he`, `hi`, `id`, `pl`, `pt-BR`, `tr`, `zh-CN`.
BCP-47 tags use camelCase identifiers (`ptBR`, `zhCN`).

Each locale also ships as its own subpath (`@betteroffice/pptx-i18n/pl`) so an
app that picks the locale at runtime can code-split rather than bundle them all.
For lookup by tag, `import { locales } from '@betteroffice/pptx-i18n'` (pulls
every locale into the bundle).

## Types and helpers

The package exports `LocaleStrings`, `PartialLocaleStrings`, `Translations`,
`TranslationKey`, `LocaleCode`, `TFunction`, `createT`, and `deepMerge`. For
non-React hosts, build a typed `t()` directly:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/pptx-i18n';

const t = createT(deepMerge(en, de) as LocaleStrings, 'de');
t('toolbar.addSlide');
```

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/pptx-i18n/CHANGELOG.md) · Apache-2.0.
