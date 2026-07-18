# @betteroffice/docx-i18n

Shared locale strings, types, and runtime helpers for the [OpenOOXML DOCX editor](https://betteroffice.dev) adapters. One source of truth for translations consumed by `@betteroffice/docx-react` and `@betteroffice/docx-vue`.

## Quick Start

```bash
npm install @betteroffice/docx-i18n
```

Pass a typed locale to the editor's `i18n` prop:

```tsx
// React
import { de } from '@betteroffice/docx-i18n';
<DocxEditor documentBuffer={file} i18n={de} />

// Vue
import { de } from '@betteroffice/docx-i18n';
<DocxEditor :document-buffer="file" :i18n="de" />
```

Mix a community locale with custom overrides:

```ts
import { de } from '@betteroffice/docx-i18n';

const myLocale = {
  ...de,
  toolbar: { ...de.toolbar, bold: 'Fettdruck' },
};
```

Keys set to `null` in any locale fall back to English.

## Packages

| Package                                                                          | Description                                                                                                                                |
| -------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| [`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react)   | <img src="https://cdn.simpleicons.org/react/61DAFB" width="20" align="middle" /> &nbsp; React adapter. Toolbar, paged editor, plugins.     |
| [`@betteroffice/docx-vue`](https://www.npmjs.com/package/@betteroffice/docx-vue)       | <img src="https://cdn.simpleicons.org/vuedotjs/4FC08D" width="20" align="middle" /> &nbsp; Vue 3 adapter. Toolbar, paged editor, plugins.  |
| [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx)               | Framework-agnostic core: OOXML parser, serializer, layout engine, ProseMirror schema. Depend on this if you fork the React or Vue adapter. |
| [`@betteroffice/docx-i18n`](https://www.npmjs.com/package/@betteroffice/docx-i18n)     | Shared locale strings and types consumed by both adapters.                                                                                 |

> **Forking the adapter?** Keep your fork thin. Depend on `@betteroffice/docx` directly so parser, serializer, and rendering fixes land in your build automatically, without backporting each upstream change by hand.

## Available locales

| Code    | Export | Language            |
| ------- | ------ | ------------------- |
| `en`    | `en`   | English (source)    |
| `de`    | `de`   | German              |
| `he`    | `he`   | Hebrew              |
| `pl`    | `pl`   | Polish              |
| `pt-BR` | `ptBR` | Portuguese (Brazil) |
| `tr`    | `tr`   | Turkish             |
| `zh-CN` | `zhCN` | Simplified Chinese  |

BCP-47 codes (`pt-BR`, `zh-CN`) use camelCase JS identifiers (`ptBR`, `zhCN`). For runtime lookup by tag:

```ts
import { locales } from '@betteroffice/docx-i18n';
<DocxEditor i18n={locales[userPreferredLocale]} />
```

> Importing `locales` pulls every locale into your bundle. For a smaller bundle, import only the ones you need by name; `sideEffects: false` lets the rest tree-shake.

## Per-locale subpaths

For apps that pick the locale at runtime, the named exports above don't tree-shake — the bundler can't know which locale wins, so it ships them all. Use the per-locale subpaths instead. Each one bundles a single locale (~30KB) and code-splits cleanly:

```ts
// Static — bundler ships only this locale's strings
import pl from '@betteroffice/docx-i18n/pl';

// Dynamic — splits into its own chunk, loaded on demand
const pl = (await import('@betteroffice/docx-i18n/pl')).default;
```

Subpaths ship for every locale: `/en`, `/de`, `/he`, `/pl`, `/pt-BR`, `/tr`, `/zh-CN`. Each also exports its locale as a named binding (`import { pl } from '@betteroffice/docx-i18n/pl'`) for callers that prefer non-default imports.

## Types

```ts
import type {
  LocaleStrings, // shape of `en`, the full source of truth
  PartialLocaleStrings, // shape of a community partial (null falls back)
  Translations, // alias for PartialLocaleStrings
  TranslationKey, // 'toolbar.bold' | 'dialogs.findReplace.title' | ...
  LocaleCode, // 'en' | 'de' | 'pt-BR' | ...
  TFunction, // signature of the `t()` callback
} from '@betteroffice/docx-i18n';
```

## Non-React/Vue hosts

Build a typed `t()` outside the adapter packages:

```ts
import { createT, deepMerge, en, de, type LocaleStrings } from '@betteroffice/docx-i18n';

const merged = deepMerge(en, de) as LocaleStrings;
const t = createT(merged, 'de');
t('toolbar.bold'); // 'Fett'
t('dialogs.findReplace.matchCount', { current: 3, total: 15 }); // ICU plurals
```

`en.json` is the source of truth. Add keys there, then run `bun run i18n:fix` from the repo root to sync community locales (new keys land as `null`). Full guide: [docs/i18n.md](https://github.com/openooxml/betteroffice/blob/main/docs/i18n.md).

## Contributing

Contributions welcome. See [CONTRIBUTING.md](https://github.com/openooxml/betteroffice/blob/main/CONTRIBUTING.md) for setup, tests, and the one-time CLA signature.

## Commercial Support

> [!TIP]
> Questions or feature requests? Open an issue at **[github.com/openooxml/betteroffice](https://github.com/openooxml/betteroffice/issues)**.
