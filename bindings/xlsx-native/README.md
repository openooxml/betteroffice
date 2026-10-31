# @betteroffice/xlsx-native

The `@betteroffice/xlsx` API on the native Rust engine, for Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

```sh
npm install @betteroffice/xlsx-native
```

This package has the same entry points (`.`, `./headless`, `./worker`, `./collaboration`), exports
and type declarations as `@betteroffice/xlsx`, so code written for one runs on the other; pick the
engine by the package you import. `initWasm()` needs no WebAssembly input here and ignores any it is given.

```ts
import { initWasm, openWorkbook } from '@betteroffice/xlsx-native';

export async function updateWorkbook(bytes: Uint8Array): Promise<Uint8Array> {
  await initWasm();
  const workbook = openWorkbook(bytes);
  try {
    workbook.editCell(0, 9, 2, '=SUM(C1:C9)');
    return workbook.save();
  } finally {
    workbook.dispose();
  }
}
```

Browsers need the WebAssembly engine: the `browser` export condition of this package throws, so
use `@betteroffice/xlsx` there. Main-thread edit peers hydrated from a worker session
(`retainPeerHydration`) also need the WebAssembly engine.
