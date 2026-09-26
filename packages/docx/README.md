# @betteroffice/docx

Framework-free core for the BetterOffice DOCX editor — the Rust engine (OOXML
parse/serialize, CRDT editing core, text shaping, pagination) compiled to
WebAssembly, plus the display-list, canvas-render, geometry, and accessibility
helpers the adapters build on. Layout never touches the DOM: the engine
measures every line and pages are replayed onto canvas.

```bash
bun add @betteroffice/docx
```

Most apps want the turnkey React component in
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react).
Reach for this package directly for headless parsing/serialization or when
building a custom adapter.

## Parse and save

A full round trip: open a `.docx`, inspect the typed model, write bytes back.

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { parseDocx, repackDocx } from '@betteroffice/docx/docx';

const document = await parseDocx(await readFile('contract.docx'));
// document.package: body, styles, numbering, theme, media, headers/footers

const bytes = await repackDocx(document);
await writeFile('contract-out.docx', Buffer.from(bytes));
```

`repackDocx` round-trips against the original buffer so untouched parts are
preserved; use `createDocx` for documents built from scratch.

The engine ships as four wasm assets (container, parser, layout, editing core)
in `dist/generated/`. Browsers fetch them lazily behind the async entry points
(`parseDocx`, save, the layout engine, `createYrsSession`); Node and Bun read
them from disk synchronously on first use. No manual init call is required.

## Collaboration

Connect the editor's Yrs replica to any reliable binary transport:

```ts
import { CollaborationProvider } from '@betteroffice/docx/collaboration';

const provider = new CollaborationProvider(replica, createTransport(), {
  user: { name: "Ada" }, // identity for this peer's remote caret
});
provider.connect();
```

`replica` can be a direct `YrsSession`, the worker-aware adapter returned by
`createWorkerCollaborationReplica`, or the value published by the React
editor's `collaboration.onReplica` callback. The provider speaks Yjs sync-v1;
room routing, authentication, awareness, and reconnection policy remain
transport concerns. Pass a persisted Yrs update as `collaboration.initialUpdate`
when a React editor joins an existing room so it hydrates the shared history
instead of independently importing the same DOCX.

## Development

The generated `.wasm` binaries are intentionally not committed. From the
repository root, install `wasm-pack` 0.15.0 and `binaryen`, then run
`bun run build:docx-wasm`.
Package builds, demo startup, and CI run this step automatically.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/docx/CHANGELOG.md) · Apache-2.0.

### Undo capture modes and boundaries

`session.setUndoCaptureMode(mode)` selects how tracked local transactions are
combined into undo steps. `session.undoCaptureMode()` returns the current mode.

| Mode | Grouping |
| --- | --- |
| `auto` (default) | Edits within 500 ms coalesce; switching stories closes the group. |
| `manual` | Edits coalesce across pauses and stories until an explicit boundary. |

`session.addUndoBoundary()` closes the current group in any mode. Changing modes
also closes the group; setting the same mode again does not. Both operations
preserve undo/redo history and are safe before capture starts. Repeated boundaries
create no empty undo steps. Remote transactions remain outside local undo history.

```ts
session.setUndoCaptureMode('manual');
session.insertText(firstLocation, 'Prefixo ');
session.insertText(secondLocation, 'Suffix');
session.addUndoBoundary();
session.deleteRange(markerRange);
session.addUndoBoundary();
session.setUndoCaptureMode('auto');
```

The two insertions undo together; marker deletion is a separate step. Boundaries
are also available in automatic mode when a host action needs to be isolated from
surrounding typing. Undo/redo close capture as usual. Manual mode controls history
grouping; it does not defer updates, flush pending input, or provide atomic execution.
The host must close manual groups so later unrelated edits do not join them.

### Paragraph identities

A session addresses paragraphs by session keys (`YrsLoc.paraId`); Word stores its
own paragraph ID (`w14:paraId`) in the file. A session key is never saved; a session
anchor resolves on every replica of one collaborative session, and each seeding open
(`openDocx`, `seedFromDocx`, `documentToYrs`) starts a new one unless given a fixed
`generation`, as a deterministic shared seed needs. Source IDs are kept as authored, and every
paragraph authored in the session gets a fresh, valid ID that avoids every ID the
package already uses.

`saveYrsDocx(session)` saves a session opened from DOCX bytes and returns each saved
paragraph's persisted anchor, qualified by its package part:

```ts
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx/yrs';

const { secondParaId } = session.splitParagraph({ story: 'body', paraId, offset: 12 });
const saved = await saveYrsDocx(session);
const anchor = saved.paragraphs.find((p) => p.session.paraId === secondParaId)!.persisted;

const reopened = await createYrsSession();
reopened.openDocx(saved.bytes, true);
reopened.resolveParagraphAnchor(anchor); // { status: 'found', anchor: { kind: 'session', … } }
```

`saved.conflicts` lists saved paragraphs whose ID the live session reassigned while the
save ran, such as by a duplicate repair after a remote update; their anchors find them in
the saved bytes only. Source paragraphs without an ID save without one and have no
persisted anchor.
`session.persistParagraphIds()` assigns them IDs across every story part, comments,
note separators and retained XML included, and repairs duplicates: source IDs and
saved IDs keep theirs over copies. It refuses, changing nothing, rather than guess at
an ambiguous comment reference. The change is replicated, stays out of undo history,
and later saves keep it; a save whose stories are otherwise unchanged patches the IDs
into the source bytes. `session.paragraphIdentities()` lists each paragraph's session,
persisted and exact-source anchors. A persisted anchor is scoped to the document the
host chose, repeated source IDs resolve as `ambiguous`, and table, cell and
content-control identities are not persisted.

### Version-checked edit batches

Read what you will target together with the session version, then apply a batch
against that version. Every step resolves against the state that was read; the
batch commits as one transaction or returns a typed refusal with nothing changed.

```ts
const read = session.readParagraphs({ story: 'body', view: 'accepted' });
if (!read.ok) throw new Error(read.failure.message);
const clause = read.paragraphs.find((paragraph) => paragraph.text.startsWith('Term'))!;

const result = session.applyEdits({
  expectVersion: read.version,
  steps: [
    {
      op: 'replaceText',
      target: { kind: 'search', text: '30 days', view: 'accepted',
        within: { kind: 'paragraph', story: 'body', paraId: clause.paraId } },
      text: '45 days',
      expect: { text: '30 days' },
    },
    {
      op: 'insertParagraphs',
      target: { story: 'body', paraId: clause.paraId },
      at: 'end',
      paragraphs: [{ text: 'Renewal is automatic.' }],
    },
  ],
});
if (!result.ok) console.warn(result.failure.code, result.failure.message);
```

`version()` changes with every committed change, local or remote (including
undo and redo), and when the session reopens its document. Tokens are scoped to
one session; after a lost response, read again rather than retrying with a new
version.

Offsets are UTF-16 positions into a paragraph's projected text. Each inline atom
(hard break, image, content control, note reference, field, other embed) is one
U+FFFC listed in `atoms`, tabs stay `\t`, and paragraph marks are excluded. The
`accepted` view includes pending insertions and hides pending deletions;
`original` does the reverse. `findText` is exact, case-sensitive and
paragraph-local, and a `search` target must match exactly once in its scope.

| Step | Effect |
| --- | --- |
| `insertText` | Inserts at the start or end of a target, formatted like typing there. |
| `replaceText` | Replaces a target's text inside one paragraph, keeping the paragraph. |
| `deleteText` | Deletes a target's text inside one paragraph. |
| `insertParagraphs` | Inserts complete paragraphs before or after an anchor, in the given or the anchor's style. |
| `deleteParagraphs` | Deletes a contiguous span of complete paragraphs without merging properties into a neighbour. |
| `setParagraphStyle` | Applies a document paragraph style with its paragraph and run formatting. |

Refusals carry a `code` (`stale-version`, `missing-target`, `ambiguous-target`,
`content-mismatch`, `overlapping-steps`, `locked-target`,
`tracked-revision-conflict`, `unsupported`, `invalid-step`, `limit-exceeded`),
the failing `stepIndex`, and the target. Editor hosts add `read-only` for an
editor that does not accept edits. Malformed requests throw.
`validateEdits(request)` runs the same checks and previews each step without
changing anything or reserving ids.

An applied batch is exactly one undo step in both capture modes; pass
`history: 'none'` to keep it out of undo history while preserving existing undo
and redo entries. `source: 'agent'` records provenance only. Add
`suggest: { author, date }` to a text step to record it as a tracked change.

Current limits: text targets stay within one paragraph; inline atoms cannot be
replaced or deleted; content-locked controls, existing tracked changes a step
would touch (including tracked run formatting), and pending paragraph-mark
revisions refuse. Paragraph and style steps need a session opened from DOCX
bytes (`openDocx`/`seedFromDocx`) and cannot be suggested. List numbering is a
v1 limitation: restyling a numbered paragraph, or applying or inserting a style
that defines numbering, refuses with `unsupported` until a follow-up retains
numbering definitions in the session. Deleting spans that hold tables,
controls, section breaks, opaque XML, fields' cached results, or comments and
bookmarks crossing the span refuses, as does any step after which saving would
move an opaque XML block, such as one that precedes a table. A batch holds at most 128 steps,
1,048,576 inserted UTF-16 units and 1,024 new paragraphs. Paragraph ids are
session anchors: they are not guaranteed to survive save and reopen.
