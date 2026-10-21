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
| `openDocx(bytes, { name?, renderer? })` | Opens a DOCX file in its own session. |
| `attachDocx(session, options?)` | Wraps a live editor's `YrsSession`. Closing the adapter leaves the session open. |
| `openXlsx(bytes, { name?, wasm? })` | Opens an XLSX workbook: cell reads and cell proposals. |
| `attachXlsx(handle, options?)` | Wraps a live `WorkbookHandle`. |
| `openPptx(bytes, { name?, wasm? })` | Opens a PPTX presentation: outline, slide reads, text proposals, atomic `edit` batches and `preview`. |
| `attachPptx(handle, options?)` | Wraps a live `PresentationHandle`. |
| `renderDocxPage` (from `@betteroffice/agents/render`) | Node canvas renderer for DOCX page previews. |
| `renderPptxSlide` | Node canvas renderer for PPTX slide previews. |
| `createOfficeMcpServer({ root, renderer?, readOnly? })` (from `@betteroffice/agents/mcp`) | The MCP server, for hosts that connect their own transport. |

DOCX and PPTX documents stage text edits with `grep`, `read` and `propose`; workbooks use
`readCells` and `proposeCells` instead. All adapters share `review`, `accept`, `reject`, `verify`,
`export` and `close`. `openXlsx` and `openPptx` load the bundled WebAssembly in Node; in a
browser, pass the module bytes as `wasm`. DOCX previews in a browser need a custom
`DocumentRenderer` instead of the Node canvas renderer.

## MCP tools

| tool | purpose |
| --- | --- |
| `office_files` | List DOCX, XLSX and PPTX files, subdirectories and open document IDs. |
| `office_open` | Open a file and return its document ID, format and capabilities. |
| `office_outline` | Page through DOCX paragraphs, PPTX slide paragraphs or XLSX sheets. |
| `office_grep` | Literal search with context; returns match IDs or cell handles. |
| `office_read` | Read a paragraph or a cell field; DOCX reads include run formatting. |
| `office_cells` | XLSX: read an A1 range with values, formulas and cell handles. |
| `office_propose` | DOCX/PPTX: stage 1-32 text replacements as `{ match, newText }`. |
| `office_propose_cells` | XLSX: stage 1-32 whole-cell inputs or formulas. |
| `office_review` | List proposals or show one proposal's before/after changes. |
| `office_accept` | Apply a proposal as one undo step; DOCX can keep it as tracked changes. |
| `office_reject` | Discard a pending proposal. |
| `office_render` | DOCX: PNG of a current or proposed page. |
| `office_verify` | Save and reopen in memory; reports which checks ran. |
| `office_export` | Write accepted edits or a pending proposal to a new file. |
| `office_close` | Release an open document and its proposals. |
| `docx_edit` | DOCX: atomic direct edits of paragraph styles, bold/italic/underline, paragraphs and table rows; `runFormat: 'anchor'` copies the first plain run on insertion. |
| `xlsx_outline`, `xlsx_read_range` | XLSX: sheets, used ranges, tables and defined names; range reads with a version. |
| `xlsx_write_range`, `xlsx_format`, `xlsx_sort` | XLSX: versioned value and formula writes, formatting and sorting. |
| `xlsx_edit_grid`, `xlsx_edit_sheets` | XLSX: insert or delete rows and columns; add, rename, delete or move sheets. |
| `xlsx_find_replace`, `xlsx_preview` | XLSX: find and replace displayed text; PNG of a range. |
| `pptx_outline`, `pptx_read_slide` | PPTX: slides, layouts and shapes; shape tree, paragraphs, runs and notes. |
| `pptx_edit`, `pptx_preview` | PPTX: atomic text, slide, shape, image and notes edits; PNG of a slide. |

## Edits and references

- References (`ref`, `match`, `revision`, `version`) belong to the open session. Reopening a file
  creates new ones. Page, slide and paragraph numbers are one-based; text offsets are zero-based.
- DOCX and PPTX search is literal and case-insensitive by default. XLSX `office_grep` searches
  displayed values case-sensitively.
- A proposal holds 1-32 separate replacements and is validated on a temporary copy before one
  atomic update. A changed target makes it stale (`STALE_TARGET`): read again and propose again.
- New text takes the formatting of the first replaced character. To keep several styles in a
  paragraph, replace the individual spans instead of the whole paragraph.
- Paragraph breaks, inline embeds and existing tracked changes are protected from text proposals.
- A committed XLSX or PPTX edit invalidates all outstanding handles, cursors and proposals for that
  file.
- Results are paginated (`nextCursor`, `nextOffset`, `nextStart`). A page of results is never an
  exhaustive search.

## Limits

- Reads return at most 16,000 UTF-16 units and 100 formatting runs; lists return at most 100 items.
- The MCP server opens at most 10 files of up to 64 MiB each and keeps at most 64 proposals per
  document. All state is in memory: export before closing the client.
- MCP messages over stdin are capped at 1 MiB.
- `verify` confirms that the saved file reopens. It does not check visual fidelity or whether the
  edit is semantically correct; use reads and previews for that.
- Previews use the BetterOffice layout engine and bundled fonts and are not pixel-identical to
  Office. Install `@betteroffice/fonts-cjk` for CJK glyphs.

## Documentation

- [Agents guide](https://docs.betteroffice.dev/docs/agents)
- [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx)
- [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx)
- [`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx)
- [Format eval suite](https://github.com/openooxml/betteroffice/blob/main/eval/README.md)
- [Repository](https://github.com/openooxml/betteroffice/tree/main/packages/agents)
