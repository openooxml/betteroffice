# @betteroffice/pptx-native

The `@betteroffice/pptx` API on the native Rust engine, for Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU and musl Linux arm64/x64, and Windows arm64/x64.

Worker sessions on Node.js >= 20.16 / 22.3 run on `worker_threads`, keeping the process alive until the session is disposed.

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

Default fonts need no setup: the Latin, international and CJK font packs install with this
package and each face is read from disk the first time a slide needs it; nothing is fetched.

Browsers need the WebAssembly engine: the `browser` export condition of this package throws, so
use `@betteroffice/pptx` there.

## Memory and untrusted input

The native engine runs inside the calling process. In `@betteroffice/pptx` each engine runs in a
WebAssembly memory of at most 4 GiB: a document that needs more makes that engine fail (worker
sessions report it as an error, in some APIs a typed out-of-memory error) while the process keeps
running. The native engine has no such cap: a document that needs more memory than the machine has
makes the whole process run out of memory.

Both engines bound their input the same way. A package is refused above 512 MiB uncompressed or
5,000 parts, and XML nested deeper than 256 levels is refused. Rendering refuses decks that need more than 256 font faces.

When you open documents from untrusted sources with this package on a server, run it under a
memory limit (a container or cgroup limit, or a separate process) so an oversized document cannot
take the server down.
