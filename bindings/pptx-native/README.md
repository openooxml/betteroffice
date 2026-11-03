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

## Native vs WebAssembly

Choose native for servers and CLIs in Node, Bun or Deno where a prebuilt binary exists; use WebAssembly in browsers or where an addon can't load.

Native is 14 to 47% faster for most operations and uses 8 to 35% less memory at peak. Edits and export on the 14 MB deck show no clear difference. Engine start-up takes about 11 ms instead of 50 ms.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 150-slide deck | Open | 274 ms | 228 ms | 0.85 (0.71 to 0.94) |
| 150-slide deck | Edits | 69.3 ms | 56.7 ms | 0.82 (0.77 to 0.88) |
| 150-slide deck | Save | 76.4 ms | 62.6 ms | 0.81 (0.70 to 0.86) |
| 150-slide deck | Layout | 304 ms | 242 ms | 0.79 (0.71 to 0.88) |
| 150-slide deck | Export | 372 ms | 323 ms | 0.86 (0.77 to 0.90) |
| 150-slide deck | Peak memory | 342 MiB | 313 MiB | 0.92 (0.91 to 0.93) |
| 2 MB government deck | Open | 41.8 ms | 23.5 ms | 0.53 (0.51 to 0.56) |
| 2 MB government deck | Edits | 4.1 ms | 2.9 ms | 0.68 (0.58 to 0.83) |
| 2 MB government deck | Save | 9.8 ms | 5.2 ms | 0.54 (0.48 to 0.65) |
| 2 MB government deck | Layout | 26.3 ms | 20.5 ms | 0.70 (0.67 to 0.79) |
| 2 MB government deck | Export | 49.5 ms | 26.9 ms | 0.56 (0.50 to 0.63) |
| 2 MB government deck | Peak memory | 204 MiB | 133 MiB | 0.65 (0.62 to 0.67) |
| 14 MB government deck | Open | 209 ms | 186 ms | 0.81 (0.70 to 0.89) |
| 14 MB government deck | Edits | 12.8 ms | 10.3 ms | 0.82 (0.71 to 1.01, no clear difference) |
| 14 MB government deck | Save | 55.9 ms | 33.5 ms | 0.56 (0.53 to 0.60) |
| 14 MB government deck | Layout | 59.9 ms | 46.7 ms | 0.72 (0.68 to 0.79) |
| 14 MB government deck | Export | 226 ms | 221 ms | 0.98 (0.94 to 1.02, no clear difference) |
| 14 MB government deck | Peak memory | 579 MiB | 502 MiB | 0.87 (0.80 to 0.88) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are a text box and 100 text inserts, layout is every slide, and export is Markdown. Timings varied by 15 to 30% between rounds, so the ratios are steadier than the absolute values.

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
