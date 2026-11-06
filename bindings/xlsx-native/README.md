# @betteroffice/xlsx-native

The `@betteroffice/xlsx` API on the native Rust engine, for Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU and musl Linux arm64/x64, and Windows arm64/x64.
PNG export renders with the installed Latin, international and CJK font packs; nothing is fetched.

Worker sessions on Node.js >= 20.16 / 22.3 run on `worker_threads`, keeping the process alive until the session is disposed.

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

## Native vs WebAssembly

Choose native for servers and CLIs in Node, Bun or Deno where a prebuilt binary exists; use WebAssembly in browsers or where an addon can't load.

Native is faster for every operation, with edits about three times as fast. Its peak memory is 16 to 18% higher on the 50,000-row and 4.1 MB workbooks and 5% lower on the 0.8 MB workbook. Engine start-up takes about 12 ms instead of 50 ms.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 50,000-row workbook with 50,000 formulas | Open | 1,937 ms | 1,405 ms | 0.72 (0.67 to 0.79) |
| 50,000-row workbook with 50,000 formulas | Edits | 293 ms | 98.1 ms | 0.33 (0.32 to 0.34) |
| 50,000-row workbook with 50,000 formulas | Save | 579 ms | 392 ms | 0.67 (0.65 to 0.70) |
| 50,000-row workbook with 50,000 formulas | Layout | 19.0 ms | 12.8 ms | 0.67 (0.64 to 0.71) |
| 50,000-row workbook with 50,000 formulas | Export | 1,977 ms | 1,515 ms | 0.74 (0.71 to 0.82) |
| 50,000-row workbook with 50,000 formulas | Peak memory | 607 MiB | 700 MiB | 1.16 (1.14 to 1.16) |
| 0.8 MB government workbook | Open | 717 ms | 513 ms | 0.67 (0.64 to 0.73) |
| 0.8 MB government workbook | Edits | 4,331 ms | 1,325 ms | 0.30 (0.28 to 0.34) |
| 0.8 MB government workbook | Save | 66.4 ms | 44.5 ms | 0.64 (0.61 to 0.75) |
| 0.8 MB government workbook | Layout | 14.3 ms | 8.5 ms | 0.60 (0.57 to 0.63) |
| 0.8 MB government workbook | Export | 784 ms | 563 ms | 0.70 (0.63 to 0.72) |
| 0.8 MB government workbook | Peak memory | 439 MiB | 418 MiB | 0.95 (0.91 to 0.96) |
| 4.1 MB government statistics workbook | Open | 5,250 ms | 1,889 ms | 0.36 (0.35 to 0.40) |
| 4.1 MB government statistics workbook | Edits | 1,316 ms | 361 ms | 0.28 (0.27 to 0.29) |
| 4.1 MB government statistics workbook | Save | 172 ms | 106 ms | 0.63 (0.60 to 0.68) |
| 4.1 MB government statistics workbook | Layout | 12.0 ms | 7.0 ms | 0.59 (0.58 to 0.62) |
| 4.1 MB government statistics workbook | Export | 5,460 ms | 2,029 ms | 0.38 (0.37 to 0.42) |
| 4.1 MB government statistics workbook | Peak memory | 983 MiB | 1,165 MiB | 1.18 (1.16 to 1.22) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are 200 cell edits, layout is 20 viewports of display list, and export is Markdown. Timings varied by up to about 25% between rounds, so the ratios are steadier than the absolute values.

## Memory and untrusted input

The native engine runs inside the calling process. In `@betteroffice/xlsx` each engine runs in a
WebAssembly memory of at most 4 GiB; the native engine gives each thread the same 4 GiB budget. It
counts the engine's live allocations on that thread, and a call that takes them past the budget
fails: a worker session fails with a typed out-of-memory error, a workbook handle throws and then
refuses every call except `dispose()`, and the process keeps running.

Both engines bound their input the same way. A package is refused above 512 MiB uncompressed or
5,000 parts, and XML nested deeper than 256 levels is refused. Worksheets are read up to 10,000,000 cells and the shared string table up to 10,000,000 entries.

The budget counts live engine allocations, not process memory, and allocations are checked as the
engine works, so process memory can exceed it: in a measurement with the budget lowered to 1 GiB,
opening a 1,000,000-row workbook peaked about 40% above the budget, from allocator overhead and the
JavaScript side. When you open
documents from untrusted sources with this package on a server, set the container or cgroup memory
limit above the budget of every thread that opens documents, with headroom for that overshoot, or
open them in a separate process.
