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

Native is 7 to 48% faster for most operations and uses 8 to 34% less memory at peak. Export on the 14 MB deck shows no clear difference. Engine start-up takes about 11 ms instead of 47 ms.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 150-slide deck | Open | 252 ms | 218 ms | 0.83 (0.77 to 0.91) |
| 150-slide deck | Edits | 67.2 ms | 54.2 ms | 0.82 (0.73 to 0.86) |
| 150-slide deck | Save | 74.0 ms | 58.0 ms | 0.78 (0.74 to 0.88) |
| 150-slide deck | Layout | 308 ms | 229 ms | 0.73 (0.70 to 0.75) |
| 150-slide deck | Export | 359 ms | 290 ms | 0.80 (0.72 to 0.82) |
| 150-slide deck | Peak memory | 341 MiB | 313 MiB | 0.92 (0.91 to 0.93) |
| 2 MB government deck | Open | 41.3 ms | 24.1 ms | 0.59 (0.52 to 0.60) |
| 2 MB government deck | Edits | 4.5 ms | 2.6 ms | 0.68 (0.57 to 0.73) |
| 2 MB government deck | Save | 10.1 ms | 5.1 ms | 0.52 (0.48 to 0.55) |
| 2 MB government deck | Layout | 26.4 ms | 19.2 ms | 0.73 (0.65 to 0.76) |
| 2 MB government deck | Export | 49.2 ms | 26.9 ms | 0.55 (0.53 to 0.63) |
| 2 MB government deck | Peak memory | 204 MiB | 133 MiB | 0.66 (0.59 to 0.67) |
| 14 MB government deck | Open | 218 ms | 197 ms | 0.93 (0.64 to 0.98) |
| 14 MB government deck | Edits | 12.8 ms | 11.0 ms | 0.82 (0.78 to 0.97) |
| 14 MB government deck | Save | 61.3 ms | 33.1 ms | 0.53 (0.51 to 0.60) |
| 14 MB government deck | Layout | 63.3 ms | 48.0 ms | 0.74 (0.68 to 0.84) |
| 14 MB government deck | Export | 230 ms | 221 ms | 0.97 (0.91 to 1.03, no clear difference) |
| 14 MB government deck | Peak memory | 582 MiB | 475 MiB | 0.81 (0.80 to 0.85) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are a text box and 100 text inserts, layout is every slide, and export is Markdown. Timings varied by up to about 25% between rounds, so the ratios are steadier than the absolute values.

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
