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

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/docx`.
CommonJS `require()` is also supported.
