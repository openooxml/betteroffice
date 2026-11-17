# @betteroffice/agents

Search, read, edit and preview DOCX, XLSX and PPTX files from an agent, through a TypeScript SDK
or the `betteroffice-mcp` server. Agents work on large files the way they work on source code:
grep, read a small region, make an exact edit, check the result, export.

No model provider, API key or agent framework is required.

## Install

```bash
bun add @betteroffice/agents
```

Requires Node.js 22 or later.

## Usage

### MCP server

```bash
bunx --package @betteroffice/agents betteroffice-mcp --root /absolute/path/to/documents
```

The server speaks MCP over stdin/stdout, so connect it from an MCP client. For example, with Codex:

```bash
codex mcp add betteroffice -- /absolute/path/to/node_modules/.bin/betteroffice-mcp --root /absolute/path/to/documents
```

| flag | description |
| --- | --- |
| `--root <directory>` | Required. Every file path is resolved against it, including symlinks. |
| `--read-only` | Exposes only the inspection tools: no proposals, edits, acceptance or export. |
| `--max-edits <n>` | Maximum edits in `docx_edit` and `office_propose`: default 32, range 1-500. |
| `--formats <list>` | Comma-separated formats to serve (`docx`, `xlsx`, `pptx`); env `BETTEROFFICE_MCP_FORMATS`. The flag wins over the env var. |
| `--tools <list>` | Comma-separated tool names to expose; unknown names fail at startup. Env `BETTEROFFICE_MCP_TOOLS`. The flag wins over the env var. |

Exports always write a new file and never overwrite an existing one. Path checks are not an
operating-system sandbox, so use a workspace directory you control.

### SDK

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
  const proposal = doc.propose({
    author: 'report-agent',
    note: 'Update the executive summary revenue',
    edits: [{ match: hit.match, newText: '€5.1 million' }],
  });

  const preview = await doc.render(1, proposal.id);
  await writeFile('preview.png', preview.png);
  console.log(doc.review(proposal.id), await doc.verify(proposal.id));
  await writeFile('report-revised.docx', await doc.export(proposal.id), { flag: 'wx' });
} finally {
  doc.close();
}
```

Exporting a pending proposal leaves the document unchanged. To apply it, call
`await doc.accept(proposal.id)` and then `doc.export()` without an ID. For Word tracked changes,
call `doc.accept(id, { tracked: true, date: '2026-09-27T12:00:00Z' })`.

## Entry points

| export | description |
| --- | --- |
| `openDocx(bytes, { name?, renderer?, maxEdits? })` | Opens a DOCX file in its own session; `maxEdits` defaults to 32. |
| `attachDocx(session, options?)` | Wraps a live editor's `YrsSession`. Closing the adapter leaves the session open. |
| `openXlsx(bytes, { name?, wasm? })` | Opens an XLSX workbook: cell reads and cell proposals. |
| `attachXlsx(handle, options?)` | Wraps a live `WorkbookHandle`. |
| `openPptx(bytes, { name?, wasm? })` | Opens a PPTX presentation: outline, slide reads, text proposals, atomic `edit` batches and `preview`. |
| `attachPptx(handle, options?)` | Wraps a live `PresentationHandle`. |
| `renderDocxPage` (from `@betteroffice/agents/render`) | Engine PNG renderer for DOCX page previews. |
| `renderPptxSlide` | Engine PNG renderer for PPTX slide previews. |
| `createOfficeMcpServer({ root, renderer?, readOnly?, formats?, tools?, maxEdits? })` (from `@betteroffice/agents/mcp`) | The MCP server, for hosts that connect their own transport; `maxEdits` defaults to 32, range 1-500. |

DOCX and PPTX documents stage text edits with `grep`, `read` and `propose`; workbooks use
`readCells` and `proposeCells` instead. All adapters share `review`, `accept`, `reject`, `verify`,
`export` and `close`. Node, Bun and Deno use native engines and their PNG exports,
with automatic WebAssembly fallback. Prebuilts cover macOS arm64/x64, Linux gnu and
musl arm64/x64, and Windows arm64/x64. `@napi-rs/canvas` is only needed for previews
on platforms without a native prebuilt. Deno needs `--allow-ffi`; without it agents
warns and uses WebAssembly. Native engines have no 4 GiB WebAssembly memory cap.
Browsers keep WebAssembly; pass module bytes as `wasm` for XLSX and PPTX.
Font defaults for DOCX previews on those runtimes are configured on `@betteroffice/docx-native/layout`.

## Attached mode

`attachXlsx(handle, { editPeer? })` and `attachDocx(session, { editPeer? })` route
proposals through the shared editor engine. Standalone `openXlsx` and `openDocx`,
including the CLI and MCP defaults, keep JS proposals. With a DOCX `editPeer`, use
`proposeAsync`, `rejectAsync` and `discardAsync`; acceptance is already async.

Agent refs, matches, cell handles and search cursors survive the agent's own
proposal operations in attached mode. DOCX editor decisions expire refs and saved
matches in touched paragraphs, including rejections that leave visible text unchanged.
Using an expired target throws `STALE_TARGET`; grep again for fresh targets. Matches
in untouched paragraphs stay valid. Editor decisions expire search cursors as user
edits do. XLSX attached proposals use engine acceptance guards, including cell styles and
dependent previews, so unrelated user edits need not stale a proposal. Stale
acceptance throws `STALE_TARGET` and leaves the proposal pending without writing
cells. Headless XLSX retains whole-workbook proposal staleness.

XLSX accept receipt diagnostic counts come from engine acceptance: `changedCells`
counts reported dependent addresses, `cycleCells` is 0, `limitedCells` counts
calculation-status addresses, `truncated` reflects the calculation budget, and
`changedSheets` counts applied target sheets.

DOCX pending proposals are live tracked changes in the shared editor session.
Editor save includes them as tracked changes; agent read, grep and default export
project the original text. Export with a proposal ID projects that proposal as
accepted. Attached acceptance settles outside undo history; tracked acceptance
retains the proposal's revision IDs, author and timestamp. Once the editor accepts
or rejects any engine record in an agent proposal group, the editor owns the rest
of that group. Agent accept, tracked accept, reject, discard and proposal-limit
eviction throw `ProposalDecidedInEditorError` with code `PROPOSAL_DECIDED_IN_EDITOR`,
without changing or withdrawing records, through both direct sessions and async
peers. `details.records` names each decided record's `id` and `state`. List and
review include per-record `records` and a derived `status`: pending while any record
is proposed, then accepted if any was accepted, otherwise rejected. Any attached accept
`date` option throws `AcceptDateUnsupportedError` with code
`ACCEPT_DATE_UNSUPPORTED`. Headless tracked acceptance still honors `date`.

Proposals over or next to existing tracked changes can throw
`TrackedRevisionConflictError` with code `TRACKED_REVISION_CONFLICT`. Its
`conflicts` list contains each conflicting `revisionId`, `author`, `kind`, `range`
and paragraph `ref` (also available in `details.conflicts`). Headless insertion
beside a tracked run retains its existing behavior.

## MCP tools

| tool | purpose |
| --- | --- |
| `office_files` | List DOCX, XLSX and PPTX files, subdirectories and open document IDs. |
| `office_open` | Open a file and return its document ID, format and capabilities. |
| `office_outline` | Page through DOCX paragraphs, PPTX slide paragraphs or XLSX sheets. |
| `office_grep` | Literal search with context; returns match IDs or cell handles. |
| `office_read` | Read a paragraph or a cell field; DOCX reads include run formatting. |
| `docx_read_runs` | DOCX: page through paragraphs with run IDs, revisions and deduplicated formats for `docx_edit` `set_runs`; fixed tokens cannot be edited. Defaults to 50 paragraphs and 32000 characters, up to 200 paragraphs per page. |
| `office_cells` | XLSX: read an A1 range with values, formulas and cell handles. |
| `office_propose` | DOCX/PPTX: stage text replacements as `{ match, newText }`, up to `maxEdits` (default 32). |
| `office_propose_cells` | XLSX: stage 1-32 whole-cell inputs or formulas. |
| `office_review` | List proposals or show one proposal's before/after changes. |
| `office_accept` | Apply a proposal as one undo step; DOCX can keep it as tracked changes. |
| `office_reject` | Discard a pending proposal. |
| `office_render` | DOCX: PNG of a current or proposed page. |
| `office_verify` | Save and reopen in memory; for DOCX also compare the saved package with the original (parts, element counts, fields, drawings, relationships). |
| `office_export` | Write accepted edits or a pending proposal to a new file. |
| `office_close` | Release an open document and its proposals. |
| `docx_edit` | DOCX: atomic direct edits of paragraph styles, bold/italic/underline, paragraphs, table rows, cell merges, comments (add, reply, resolve or reopen) and page setup with `insert_section_break` and `set_section`; `runFormat: 'anchor'` copies the first plain run on insertion. `set_runs {paragraph, revision, runs:[{run,text}]}` replaces listed runs while keeping formatting; empty text deletes. |
| `xlsx_outline`, `xlsx_read_range` | XLSX: sheets, used ranges, tables and defined names; range reads with a version. |
| `xlsx_write_range`, `xlsx_format`, `xlsx_sort` | XLSX: versioned value and formula writes, formatting and sorting. |
| `xlsx_list_defined_names`, `xlsx_edit_defined_names` | XLSX: list complete names with pagination; set or delete workbook and sheet scoped names atomically, including hidden flags. |
| `xlsx_edit_grid`, `xlsx_edit_sheets` | XLSX: insert or delete rows and columns; add, rename, delete or move sheets. |
| `xlsx_set_auto_filter` | XLSX: set, extend or remove a sheet's filter buttons; table filters are refused. |
| `xlsx_list_charts`, `xlsx_read_chart_text`, `xlsx_edit_charts` | XLSX: inspect charts on worksheets and chartsheets and page through complete titles, names and references; replace or remove titles and edit series names, category refs and value refs atomically. |
| `xlsx_list_conditional_formats`, `xlsx_read_conditional_format_formula`, `xlsx_edit_conditional_formats` | XLSX: list bounded rule previews and page through complete formulas; add or delete `cellIs` and `expression` rules with a differential format. |
| `xlsx_find_replace`, `xlsx_preview` | XLSX: find and replace displayed text; PNG of a range. |
| `pptx_outline`, `pptx_read_slide` | PPTX: slides, layouts and shapes; shape tree, paragraphs, runs and notes. `office_outline` with `story: 'comments'` lists discussions; `office_read` reads full comment text. |
| `pptx_edit`, `pptx_preview` | PPTX: atomic text, slide, layout mapping, shape, image, notes, comment (`add_comment`, `reply_comment`, `resolve_comment`) and table structure edits (`insert_table_row`, `delete_table_row`, `insert_table_column`, `delete_table_column`); PNG of a slide. |

## Edits and references

- References (`ref`, `match`, `revision`, `version`) belong to the open session. Reopening a file
  creates new ones. Page, slide and paragraph numbers are one-based; text offsets are zero-based.
- DOCX and PPTX search is literal and case-insensitive by default. XLSX `office_grep` searches
  displayed values case-sensitively.
- A proposal holds up to `maxEdits` separate replacements (default 32) and is validated on a temporary copy before one
  atomic update. A changed target makes it stale (`STALE_TARGET`): read again and propose again.
- New text takes the formatting of the first replaced character. To keep several styles in a
  paragraph, replace the individual spans instead of the whole paragraph.
- Paragraph breaks, inline embeds and existing tracked changes are protected from text proposals.
- A committed XLSX or PPTX edit invalidates all outstanding handles, cursors and proposals for that
  file.
- Results are paginated (`nextCursor`, `nextOffset`, `nextStart`). A page of results is never an
  exhaustive search.
- `readRuns` returns complete paragraphs and revision-bound run IDs. Pass them to `set_runs`;
  unlisted runs stay unchanged. Use one `set_runs` per paragraph in a batch.

## Limits

- `office_read` returns at most 16,000 UTF-16 units and 100 formatting runs; lists return at most 100 items.
- The MCP server opens at most 10 files of up to 64 MiB each and keeps at most 64 pending proposals per
  document, evicting settled proposals as needed. All state is in memory: export before closing the client.
- MCP messages over stdin are capped at 1 MiB.
- `verify` confirms that the saved file reopens. For DOCX it also compares the saved package with the
  original and lists structure the edits do not explain. It does not check visual fidelity or whether
  the edit is semantically correct; use reads and previews for that.
- Previews use the BetterOffice layout engine and bundled fonts and are not pixel-identical to
  Office. Install `@betteroffice/fonts-cjk` for CJK glyphs and `@betteroffice/fonts-intl` for additional international scripts.

## Documentation

- [Agents guide](https://docs.betteroffice.dev/docs/agents)
- [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx)
- [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx)
- [`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx)
- [Format eval suite](https://github.com/openooxml/betteroffice/blob/main/eval/README.md)
- [Repository](https://github.com/openooxml/betteroffice/tree/main/packages/agents)
