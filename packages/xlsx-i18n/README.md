# @betteroffice/xlsx-i18n

Shared locale strings, types, and runtime helpers for the [BetterOffice XLSX editor](https://betteroffice.dev) React adapter.

## Quick Start

```bash
npm install @betteroffice/xlsx-i18n
```

Pass a typed locale to the editor's `i18n` prop:

```tsx
import { de } from '@betteroffice/xlsx-i18n';
import { XlsxEditor } from '@betteroffice/xlsx-react';

<XlsxEditor file={file} i18n={de} />
```

Keys set to `null` fall back to English.

## Available locales

| Code | Export | Language |
| --- | --- | --- |
| `en` | `en` | English |
| `de` | `de` | German |
| `fr` | `fr` | French |
| `he` | `he` | Hebrew |
| `hi` | `hi` | Hindi |
| `id` | `id` | Indonesian |
| `pl` | `pl` | Polish |
| `pt-BR` | `ptBR` | Portuguese (Brazil) |
| `tr` | `tr` | Turkish |
| `zh-CN` | `zhCN` | Simplified Chinese |

For runtime lookup, import `locales`. To keep bundles smaller, import a single locale from a per-locale subpath:

```ts
import pl from '@betteroffice/xlsx-i18n/pl';
```

## Types and helpers

The package exports `LocaleStrings`, `PartialLocaleStrings`, `Translations`, `TranslationKey`, `LocaleCode`, `TFunction`, `createT`, and `deepMerge`.

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/xlsx-i18n';

const strings = deepMerge(en, de) as LocaleStrings;
const t = createT(strings, 'de');
t('toolbar.save');
```

## Commercial Support

> [!TIP]
> Questions or feature requests? Open an issue at **[github.com/openooxml/betteroffice](https://github.com/openooxml/betteroffice/issues)**.
