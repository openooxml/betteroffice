# @betteroffice/pptx-native

Native Rust bindings for the BetterOffice PPTX engine in Node.js and Bun. The package is intended for servers, CLIs, Electron applications, and other Node.js or Bun hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, collaboration behavior, and editing semantics of the Rust facade and its Python binding. Facade operations return promises and run in call order on one native worker. That also keeps the thread-affine Rust facade on its owning thread without blocking the JavaScript event loop.

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
See the [JavaScript guide](https://docs.betteroffice.dev/docs/javascript#run-natively-in-nodejs-and-bun)
for rendering and API examples. CommonJS `require()` is also supported.
