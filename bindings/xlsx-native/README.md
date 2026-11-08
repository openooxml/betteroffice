# @betteroffice/xlsx-native

The `@betteroffice/xlsx` API on the native Rust engine, for Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU and musl Linux arm64/x64, and Windows arm64/x64.
PNG export renders with the installed Latin, international and CJK font packs; nothing is fetched.

An installed font pack that cannot be used (another release, missing files, an unreadable manifest) is skipped and its
faces substituted; the process warns once per pack with code `BETTEROFFICE_FONT_PACK`, and `fontPackDiagnostics()`
lists each one with the pack and the reason.

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
