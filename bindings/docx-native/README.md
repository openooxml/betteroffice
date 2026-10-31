# @betteroffice/docx-native

The `@betteroffice/docx` API backed by native Rust in Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

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
`@betteroffice/fonts` is an optional peer for font assets.

The previous `Document` and `openDocument` facade has been removed. Use the
[DOCX API](https://docs.betteroffice.dev/docs/javascript) with the native package name.
Browser imports throw; use `@betteroffice/docx` in browsers.
