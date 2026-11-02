# @betteroffice/docx-native

The `@betteroffice/docx` API backed by native Rust in Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

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
