# @betteroffice/agents

Work with large documents the way a coding agent works with source files:
discover, grep, read a small region, propose an exact edit, inspect the result,
and export it. DOCX supports text proposals and page previews. The first
**XLSX and PPTX prototypes** add cell input/formula proposals and slide text
replacements through the same SDK and MCP server.

The SDK works with existing BetterOffice document sessions and headless files.
The local MCP server exposes the same operations to agents, including rendered
page images. No model provider, API key, or agent framework is required by this
package.

## MCP quickstart

```sh
bun add @betteroffice/agents
bunx --package @betteroffice/agents betteroffice-mcp --root /absolute/path/to/documents
```

The server speaks MCP over stdin/stdout. Connect it from an MCP client rather
than typing commands into the process. File paths are checked against `--root`,
including resolved symlinks. Exports require a new filename. Add `--read-only`
to omit proposal, acceptance, rejection, and export tools.

Use a workspace whose directory structure you control. Path checks are not an
operating-system sandbox against other local processes moving directories.
If an export fails after creating its destination, `EXPORT_INCOMPLETE` identifies
the potentially partial file and the underlying error. Inspect it and retry with
a new filename. The server does not delete the path, since another process could
have replaced it after the failure.

For Codex, register the installed executable:

```sh
codex mcp add betteroffice -- /absolute/path/to/node_modules/.bin/betteroffice-mcp --root /absolute/path/to/documents
```

Start a new Codex session after registration. Try:

> In report.docx, find the executive summary and update revenue from €4.2
> million to €5.1 million. Preserve its formatting, preview the proposed page,
> verify the result reopens, and export report-revised.docx.

## Tools

| Tool | Purpose |
| --- | --- |
| `office_files` | List DOCX, XLSX, and PPTX files, subdirectories, and open document IDs. |
| `office_open` | Open a file and discover its stories and capabilities. |
| `office_outline` | Page through paragraphs, slide paragraphs, or sheets. |
| `office_grep` | Search literal text, with context and precise refs/offsets. |
| `office_read` | Read bounded paragraph or cell text; DOCX includes run formatting. |
| `office_cells` | Read an XLSX A1 range with values, formulas, and cell handles. |
| `office_propose` | Stage exact DOCX/PPTX text replacements. |
| `office_propose_cells` | Stage XLSX whole-cell inputs or formulas. |
| `office_review` | List proposals or inspect one proposal's before/after text. |
| `office_render` | DOCX only: return a PNG of a current or proposed page. |
| `office_verify` | Save and reopen in memory; report which checks ran. |
| `office_accept` | Apply a proposal as one undo step, optionally with Word tracked changes. |
| `office_reject` | Reject a pending proposal, including one with stale targets. |
| `office_export` | Save accepted edits or a pending proposal to a new file. |
| `office_close` | Release an open document and its unsaved proposals. |

Use the `document` ID from `office_open` on subsequent calls. Use paragraph
`ref` and `revision` values returned by search or reads; these belong to the
open session. Reopening a file creates new references. `paragraph` and `page`
numbers are one-based; text offsets and pagination offsets are zero-based.
Run formatting reports `fontSizePt` and `complexScriptFontSizePt` in points,
converting Word's stored half-point values.

DOCX/PPTX search is literal and case-insensitive by default. DOCX search stays within paragraphs,
never crosses an inline embed, and excludes existing tracked changes. Follow
`nextCursor` for more search results, `nextOffset` for lists, and `nextStart`
for long paragraph reads. A document update invalidates search cursors; repeat
the query without the old cursor.

DOCX/PPTX MCP proposals accept only `{ match, newText }`. Search for the exact span you
want to replace and choose its match ID. For a repeated word, use its context
and occurrence order to select the right result. Explicit positional edits
are available in the DOCX typed SDK.

## DOCX SDK

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { openDocx } from '@betteroffice/agents';
import { renderDocxPage } from '@betteroffice/agents/render';

const doc = await openDocx(await readFile('report.docx'), {
  name: 'report.docx',
  renderer: renderDocxPage,
});

try {
  const [hit] = doc.grep({ query: '€4.2 million' }).matches;
  if (!hit) throw new Error('Revenue text not found');
  const paragraph = doc.read(hit.ref);
  const proposal = doc.propose({
    author: 'report-agent',
    note: 'Update the executive summary revenue',
    edits: [{ match: hit.match, newText: '€5.1 million' }],
  });

  const preview = await doc.render(1, proposal.id);
  await writeFile('preview.png', preview.png);
  const checks = await doc.verify(proposal.id);
  console.log(paragraph, doc.review(proposal.id), checks);
  await writeFile('report-revised.docx', await doc.export(proposal.id), { flag: 'wx' });
} finally {
  doc.close();
}
```

Exporting a pending proposal leaves the live session unchanged. To apply it,
call `await doc.accept(proposal.id)` and then `doc.export()` without an ID.
For native Word tracked changes, use
`doc.accept(id, { tracked: true, date: '2026-09-27T12:00:00Z' })`.

Attach to a live editor's authoritative `YrsSession` using
`attachDocx(session, options)`. Host edits immediately invalidate affected
references. Acceptance preserves unrelated edits and has its own undo boundary.
Closing the adapter leaves a borrowed session alive. If the host replaces the
entire document/session, create a new adapter. Pending proposals and attribution
are local to this adapter; tracked acceptance persists authorship in Word.

The DOCX adapter supports a custom `DocumentRenderer` in a browser host; `@betteroffice/agents/render` is the Node canvas adapter. A renderer
receives a disposable fork, so previews cannot alter the live editor's layout
or font state.

`createOfficeMcpServer({ root, renderer, readOnly })` from
`@betteroffice/agents/mcp` can be connected to an MCP transport by a host.
Pass `renderDocxPage` to enable PNG previews. The CLI configures it automatically.

## DOCX edit contract and shared limits

- Prefer `{ match, newText }`, copying `match` from grep. Each occurrence has
  its own ID, so repeated words can be targeted without computing offsets.
  The last 1,024 matches are retained; rerun grep if a match expires.
- Copy `revision` from a fresh read. Changes to the target's text, formatting,
  or embedded content cause `STALE_TARGET`; read and propose again. Unrelated
  edits can proceed.
- `oldText` must match exactly. If it occurs multiple times, supply `start`
  from grep. For insertion, use empty `oldText` and an explicit `start`.
- A proposal contains 1–32 separate, non-touching replacements. All offsets
  refer to the original paragraphs. The combined old/new text is limited to
  16,000 UTF-16 units. Validation and execution on a fork precede one atomic
  update of the live session.
- New text inherits the formatting of the first replaced character. Unchanged
  runs retain their formatting. To preserve several styles inside a paragraph,
  replace the individual text spans rather than the whole paragraph.
- Paragraph breaks, inline embeds, and existing tracked changes are protected.
  Reads show an embed as U+FFFC and mark protected runs. Reads expose the raw
  revision text with `ins`/`del` metadata; they are not an accepted-changes view.
  This release edits plain text, not tables, shapes, styles, or document structure.
- Reads return at most 16,000 UTF-16 units and 100 formatting runs. Search
  defaults to 20 matches and caps each response's text budget. Lists return
  at most 100 summaries. Never treat a page of results as an exhaustive search.
- The MCP server opens at most 10 files, each up to 64 MiB, and retains at most
  64 proposals per document. Closed proposal receipts may be evicted. Sessions
  and pending work are in memory; export before closing the client.

## What verification means

`verify` confirms that a generated Office file can be reopened by BetterOffice. Its
response explicitly leaves visual fidelity and semantic correctness unchecked.
Use targeted reads and before/after page renders to inspect the actual change.
The PNG renderer uses BetterOffice's Rust layout and bundled fonts, with render
warnings for unavailable images and missing glyphs. Its output is not a promise
of pixel identity with Word. Install `@betteroffice/fonts-cjk` for CJK coverage.

## XLSX and PPTX prototypes

```ts
import { openXlsx, openPptx } from '@betteroffice/agents';

const workbook = await openXlsx(await readFile('budget.xlsx'));
try {
  const sheet = workbook.list().items[0].sheetId;
  const { cells } = workbook.readCells({ sheet, range: 'B3:E3' });
  const proposal = workbook.proposeCells({
    author: 'budget-agent',
    edits: [{ cell: cells[0].cell, input: '1000' }, { cell: cells[3].cell, input: '=D3*2' }],
  });
  console.log(workbook.review(proposal.id), await workbook.verify(proposal.id));
  await writeFile('budget-revised.xlsx', await workbook.export(proposal.id), { flag: 'wx' });
} finally { workbook.close(); }

const deck = await openPptx(await readFile('slides.pptx'));
try {
  const [hit] = deck.grep({ query: '€4.2 million' }).matches;
  if (!hit) throw new Error('Revenue text not found');
  const proposal = deck.propose({
    author: 'slide-agent', edits: [{ match: hit.match, newText: '€5.1 million' }],
  });
  console.log(deck.review(proposal.id), await deck.verify(proposal.id));
  await writeFile('slides-revised.pptx', await deck.export(proposal.id), { flag: 'wx' });
} finally { deck.close(); }
```

In MCP, use `office_outline` to discover XLSX `sheetId` values, `office_cells`
with `{sheet, range: "B3:E3"}`, then `office_propose_cells` with `{cell, input}`.
Input replaces the entire cell using Excel-style input parsing; `=` starts a
formula, and empty input clears a cell. Formulas are recalculated by the engine;
its supported functions and calculation limits still apply. Read dependent
cells after acceptance, or reopen a proposed export to inspect results.
Review includes the old stored value, formula, and display text alongside the
proposed input. It does not promise that the formula calculates the intended result.

XLSX grep searches **displayed values case-sensitively**. `story` selects a sheet
ID. `office_cells` reads empty cells too, paginates in row-major order, and
returns at most 100 cells per call. Cell summaries truncate long fields; use
`office_read` with `field: "formula"`, `"value"`, or `"displayText"` and follow
`nextStart` for their full text. XLSX/PPTX searches stop at 10,000 matches and
report `truncated`; narrow the query or story when that cap is reached.

PPTX outline and grep return one-based slide numbers and paragraph refs.
Replacements stay within one paragraph; the engine refuses fields, line breaks,
or unsupported text structures that it cannot safely edit. Layout, slide order,
notes, charts, and media editing are outside this prototype.

Both adapters own their sessions; live-session attachment, PNG rendering, and
tracked changes are currently DOCX-only. `openXlsx` and `openPptx` initialize
the installed WASM assets automatically in Node; a browser host can pass
`{wasm: bytes}`. Any committed edit invalidates **all** outstanding XLSX/PPTX
handles, cursors, and proposals. Re-read and propose again; rejection remains
available for stale proposals. Accepting a batch is atomic and records one
engine undo step. Proposal metadata remains local; export keeps the original
format and does not persist proposal attribution.

## Development

From the repository root:

```sh
bun install
bun run build:fonts
bun run build:docx
bun run build:agents
bun run --filter '@betteroffice/agents' typecheck
bun test packages/agents
```

Tests use real DOCX/WASM sessions, including a 1,500-paragraph search fixture,
format-preserving export, stale proposals, tracked changes, isolated undo,
page images, and an MCP subprocess. See [the evaluation guide](eval/README.md)
for testing agent ergonomics on the installed package.
