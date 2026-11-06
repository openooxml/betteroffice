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

Native is not generally faster for documents. On the 3.3 MB report it is slower: edits, save and export take 19 to 24% longer and peak memory is 2% higher. On the two smaller documents, edits are 15 to 34% faster and peak memory is 21 to 34% lower; layout is 25% faster on the 3,000-paragraph document and save 10% faster on the 0.6 MB document, while export on the 3,000-paragraph document is 9% slower. Open shows no clear difference on any document. Engine start-up is no faster.

| Document | Operation | WebAssembly | Native | Native / WebAssembly (95% CI) |
|---|---|---:|---:|---:|
| 3,000-paragraph document | Open | 136 ms | 143 ms | 1.07 (0.97 to 1.13, no clear difference) |
| 3,000-paragraph document | Edits | 209 ms | 135 ms | 0.66 (0.64 to 0.79) |
| 3,000-paragraph document | Save | 1,034 ms | 993 ms | 0.98 (0.96 to 1.08, no clear difference) |
| 3,000-paragraph document | Layout | 1,144 ms | 842 ms | 0.75 (0.74 to 0.75) |
| 3,000-paragraph document | Export | 135 ms | 137 ms | 1.09 (1.00 to 1.23) |
| 3,000-paragraph document | Peak memory | 1,155 MiB | 907 MiB | 0.79 (0.75 to 0.80) |
| 0.6 MB government document | Open | 112 ms | 100 ms | 0.93 (0.85 to 1.08, no clear difference) |
| 0.6 MB government document | Edits | 18.3 ms | 16.0 ms | 0.85 (0.81 to 0.90) |
| 0.6 MB government document | Save | 1,081 ms | 937 ms | 0.90 (0.87 to 0.94) |
| 0.6 MB government document | Layout | 65.7 ms | 61.3 ms | 0.95 (0.92 to 1.01, no clear difference) |
| 0.6 MB government document | Export | 106 ms | 119 ms | 0.98 (0.89 to 1.25, no clear difference) |
| 0.6 MB government document | Peak memory | 479 MiB | 325 MiB | 0.66 (0.65 to 0.70) |
| 3.3 MB government report | Open | 2,177 ms | 2,091 ms | 0.95 (0.90 to 1.01, no clear difference) |
| 3.3 MB government report | Edits | 166 ms | 198 ms | 1.19 (1.01 to 1.31) |
| 3.3 MB government report | Save | 22,099 ms | 27,957 ms | 1.24 (1.18 to 1.44) |
| 3.3 MB government report | Layout | 1,107 ms | 1,203 ms | 1.08 (0.96 to 1.37, no clear difference) |
| 3.3 MB government report | Export | 1,995 ms | 2,351 ms | 1.21 (1.13 to 1.40) |
| 3.3 MB government report | Peak memory | 1,623 MiB | 1,669 MiB | 1.02 (1.02 to 1.03) |

Medians of 9 interleaved rounds on one 6-vCPU Linux machine with Node.js 24. A ratio below 1 favours native. Edits are 200 text inserts, layout is the whole document with one font, and export is Markdown. Timings varied by up to about 25% between rounds, so the ratios are steadier than the absolute values.

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
