# @betteroffice/docx-native

Native Rust bindings for the BetterOffice DOCX engine in Node.js and Bun. The package is intended for servers, CLIs, Electron applications, and other Node.js or Bun hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, and editing semantics of the Rust facade and its Python binding. Facade operations return promises and run in call order on one native worker without blocking the JavaScript event loop. Document bytes use `Buffer`, and structured contracts use ordinary JavaScript objects.

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
See the [JavaScript guide](https://docs.betteroffice.dev/docs/javascript#run-natively-in-nodejs-and-bun)
for rendering and API examples. CommonJS `require()` is also supported.
