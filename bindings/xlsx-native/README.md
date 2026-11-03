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

Native is faster for every operation, with edits about three times as fast. Its peak memory is 15 to 16% higher on the 50,000-row and 4.1 MB workbooks and 9% lower on the 0.8 MB workbook. Engine start-up takes about 10 ms instead of 55 ms.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 50,000-row workbook with 50,000 formulas | Open | 2,180 ms | 1,647 ms | 0.68 (0.66 to 0.85) |
| 50,000-row workbook with 50,000 formulas | Edits | 313 ms | 105 ms | 0.34 (0.32 to 0.38) |
| 50,000-row workbook with 50,000 formulas | Save | 611 ms | 433 ms | 0.68 (0.66 to 0.73) |
| 50,000-row workbook with 50,000 formulas | Layout | 19.4 ms | 13.2 ms | 0.71 (0.67 to 0.84) |
| 50,000-row workbook with 50,000 formulas | Export | 2,148 ms | 1,753 ms | 0.78 (0.70 to 0.99) |
| 50,000-row workbook with 50,000 formulas | Peak memory | 611 MiB | 704 MiB | 1.15 (1.14 to 1.16) |
| 0.8 MB government workbook | Open | 789 ms | 529 ms | 0.69 (0.60 to 0.77) |
| 0.8 MB government workbook | Edits | 4,683 ms | 1,411 ms | 0.31 (0.28 to 0.34) |
| 0.8 MB government workbook | Save | 67.8 ms | 49.9 ms | 0.66 (0.48 to 0.81) |
| 0.8 MB government workbook | Layout | 15.3 ms | 9.1 ms | 0.59 (0.46 to 0.66) |
| 0.8 MB government workbook | Export | 852 ms | 618 ms | 0.72 (0.61 to 0.74) |
| 0.8 MB government workbook | Peak memory | 460 MiB | 418 MiB | 0.91 (0.90 to 0.94) |
| 4.1 MB government statistics workbook | Open | 5,473 ms | 2,040 ms | 0.37 (0.35 to 0.55) |
| 4.1 MB government statistics workbook | Edits | 1,373 ms | 374 ms | 0.28 (0.27 to 0.32) |
| 4.1 MB government statistics workbook | Save | 200 ms | 119 ms | 0.65 (0.54 to 0.84) |
| 4.1 MB government statistics workbook | Layout | 13.2 ms | 7.4 ms | 0.61 (0.55 to 0.70) |
| 4.1 MB government statistics workbook | Export | 5,736 ms | 2,243 ms | 0.39 (0.38 to 0.58) |
| 4.1 MB government statistics workbook | Peak memory | 1,000 MiB | 1,164 MiB | 1.16 (1.16 to 1.19) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are 200 cell edits, layout is 20 viewports of display list, and export is Markdown. Timings varied by 15 to 30% between rounds, so the ratios are steadier than the absolute values.

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
