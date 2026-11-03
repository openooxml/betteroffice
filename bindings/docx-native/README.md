# @betteroffice/docx-native

The `@betteroffice/docx` API backed by native Rust in Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU and musl Linux arm64/x64, and Windows arm64/x64.

Worker sessions on Node.js >= 20.16 / 22.3 run on `worker_threads`, keeping the process alive until the session is disposed. Outside browsers the DOCX worker opens, reads and saves documents; for headless editing use the in-thread edit session (`openDocx` / `EditSession`), which is fully supported, because the worker editor session paints frames and needs `OffscreenCanvas`.

```sh
npm install @betteroffice/docx-native
```

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx-native/yrs';

const session = await createYrsSession();
try {
  session.seedFromDocx(await readFile('input.docx'));
  const first = session.paragraphs('body')[0];
  session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
  await writeFile('output.docx', (await saveYrsDocx(session)).bytes);
} finally {
  session.destroy();
}
```

All 28 public entry points, exports and declarations match `@betteroffice/docx`.
Select the engine by importing `@betteroffice/docx-native` or `@betteroffice/docx`;
there is no runtime engine switch. Native initialization needs no WebAssembly assets.
Default fonts need no setup: the Latin, international and CJK font packs install with this
package and each face is read from disk the first time a document needs it; nothing is fetched.

The previous `Document` and `openDocument` facade has been removed. Use the
[DOCX API](https://docs.betteroffice.dev/docs/javascript) with the native package name.
Browser imports throw; use `@betteroffice/docx` in browsers.

## Native vs WebAssembly

Choose native for servers and CLIs in Node, Bun or Deno where a prebuilt binary exists; use WebAssembly in browsers or where an addon can't load.

Native is not generally faster for documents. Open, save and export show no clear difference, and the 3.3 MB report shows no clear difference for any timing. Edits are faster on the two smaller documents, and layout on the 3,000-paragraph document. Peak memory is 23 to 35% lower on the two smaller documents and 3% higher on the report. Engine start-up is no faster.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 3,000-paragraph document | Open | 157 ms | 152 ms | 0.98 (0.93 to 1.08, no clear difference) |
| 3,000-paragraph document | Edits | 230 ms | 171 ms | 0.67 (0.63 to 0.79) |
| 3,000-paragraph document | Save | 1,122 ms | 1,149 ms | 0.97 (0.88 to 1.06, no clear difference) |
| 3,000-paragraph document | Layout | 1,211 ms | 921 ms | 0.75 (0.71 to 0.79) |
| 3,000-paragraph document | Export | 147 ms | 159 ms | 1.08 (0.96 to 1.40, no clear difference) |
| 3,000-paragraph document | Peak memory | 1,166 MiB | 895 MiB | 0.77 (0.73 to 0.79) |
| 0.6 MB government document | Open | 116 ms | 107 ms | 0.97 (0.92 to 1.18, no clear difference) |
| 0.6 MB government document | Edits | 18.5 ms | 16.1 ms | 0.87 (0.80 to 0.93) |
| 0.6 MB government document | Save | 1,076 ms | 1,039 ms | 0.97 (0.88 to 1.06, no clear difference) |
| 0.6 MB government document | Layout | 70.8 ms | 67.6 ms | 1.04 (0.81 to 1.08, no clear difference) |
| 0.6 MB government document | Export | 111 ms | 125 ms | 0.97 (0.80 to 1.25, no clear difference) |
| 0.6 MB government document | Peak memory | 494 MiB | 319 MiB | 0.65 (0.61 to 0.68) |
| 3.3 MB government report | Open | 2,394 ms | 2,213 ms | 0.91 (0.62 to 1.06, no clear difference) |
| 3.3 MB government report | Edits | 189 ms | 217 ms | 1.28 (0.50 to 1.56, no clear difference) |
| 3.3 MB government report | Save | 26,460 ms | 31,153 ms | 1.16 (0.86 to 1.30, no clear difference) |
| 3.3 MB government report | Layout | 1,321 ms | 1,191 ms | 1.01 (0.61 to 1.21, no clear difference) |
| 3.3 MB government report | Export | 2,351 ms | 2,447 ms | 1.08 (0.76 to 1.41, no clear difference) |
| 3.3 MB government report | Peak memory | 1,609 MiB | 1,669 MiB | 1.03 (1.01 to 1.06) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are 200 text inserts, layout is the whole document with one font, and export is Markdown. Timings varied by 15 to 30% between rounds, so the ratios are steadier than the absolute values.

## Memory and untrusted input

The native engine runs inside the calling process. In `@betteroffice/docx` each engine runs in a
WebAssembly memory of at most 4 GiB: a document that needs more makes that engine fail (worker
sessions report it as an error, in some APIs a typed out-of-memory error) while the process keeps
running. The native engine has no such cap: a document that needs more memory than the machine has
makes the whole process run out of memory.

Both engines bound their input the same way. A package is refused above 512 MiB uncompressed or
5,000 parts, and XML nested deeper than 256 levels is refused.

When you open documents from untrusted sources with this package on a server, run it under a
memory limit (a container or cgroup limit, or a separate process) so an oversized document cannot
take the server down.
