# @betteroffice/pptx-native

Native Rust bindings for the BetterOffice PPTX engine in Node.js, Bun and Deno.
Facade operations return promises and run in call order on one native worker.

```sh
npm install @betteroffice/pptx-native
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openPresentation } from '@betteroffice/pptx-native';

const presentation = await openPresentation(await readFile('input.pptx'));
await writeFile('output.pptx', await presentation.save());
```

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/pptx`.
CommonJS `require()` is also supported.

Edits default to author `node` and origin `local`. Set `author` and `origin` in
`openPresentation` options to supply an explicit edit context.
`setCommentStatus(commentId, resolved)` requires an explicit boolean.
