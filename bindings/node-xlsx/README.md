# @betteroffice/xlsx-native

Native Rust bindings for the BetterOffice XLSX engine in Node.js and Bun. The package is intended for servers, CLIs, Electron applications, and other Node.js or Bun hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, calculation behavior, and editing semantics of the Rust facade and its Python binding. Facade operations return promises and run in call order on one native worker without blocking the JavaScript event loop. Workbook bytes use `Buffer`.

```sh
npm install @betteroffice/xlsx-native
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@betteroffice/xlsx-native';

const workbook = await openWorkbook(await readFile('input.xlsx'));
await writeFile('output.xlsx', await workbook.save());
```

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/xlsx`.
See the [JavaScript guide](https://docs.betteroffice.dev/docs/javascript#run-natively-in-nodejs-and-bun)
for rendering and API examples. CommonJS `require()` is also supported.
