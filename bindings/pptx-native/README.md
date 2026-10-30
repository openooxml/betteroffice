# @betteroffice/pptx-native

The `@betteroffice/pptx` API on the native Rust engine, for Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

```sh
npm install @betteroffice/pptx-native
```

This package has the same entry points (`.`, `./worker`, `./collaboration`), exports
and type declarations as `@betteroffice/pptx`, so code written for one runs on the other; pick the
engine by the package you import. `initWasm()` needs no WebAssembly input here and ignores any it is given.

```ts
import { initWasm, openPresentation } from '@betteroffice/pptx-native';

export async function updatePresentation(bytes: Uint8Array): Promise<Uint8Array> {
  await initWasm();
  const presentation = openPresentation(bytes);
  try {
    const slideId = presentation.snapshot().slides[0].id;
    presentation.addTextBox(slideId, {
      name: 'Title', rect: { x: 100000, y: 200000, width: 3000000, height: 1000000 },
      text: 'Hello', style: { fontFamily: 'Arial', fontSizePt: 18 },
    });
    return presentation.save();
  } finally {
    presentation.dispose();
  }
}
```

Browsers need the WebAssembly engine: the `browser` export condition of this package throws, so
use `@betteroffice/pptx` there.
