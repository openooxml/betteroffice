# @betteroffice/docx-native

Native Rust bindings for the BetterOffice DOCX engine in Node.js, Bun and Deno.
Facade operations return promises and run in call order on one native worker.

```sh
npm install @betteroffice/docx-native
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openDocument } from '@betteroffice/docx-native';

const document = await openDocument(await readFile('input.docx'));
await writeFile('output.docx', await document.save());
```

`openDocument(bytes, { limits })` accepts partial parser limits with camelCase
names such as `maxXmlBytes`, `maxParagraphs`, and `maxTableCells`. Omitted limits
keep the Rust parser defaults.

`document.exportStructured(options)` and `document.exportMarkdown(options)`
return structured content and Markdown content objects. `revisionView` is
required (`accepted`, `original`, or `markup`); `stories`, `includeFormatting`,
`maxBlocks`, and `maxBytes` are optional. Both read the current document state.
The synchronous `renderDocxMarkdown(content, { maxBytes }?)` free function renders
structured content to a Markdown string.

`document.listContentControls(options?)` lists controls, and
`document.findContentControls(query, options?)` finds exact matches by `id`
(`controlId`), `tag`, `ooxmlId`, or `alias`. Options accept `stories`,
`maxControls`, and `maxBytes`.

Exports and control reads resolve with content directly, without a session
version or `ok` envelope. Refusals reject with an `Error` whose `message` is the
failure message and whose `failure` property contains `code`, `message`, and
`target`. The Markdown renderer throws the same error for a refusal.

`document.save({ now })` accepts `now` as an alias of `timestamp`. Supplying
different values for both rejects with a `TypeError`.

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/docx`.
CommonJS `require()` is also supported.
