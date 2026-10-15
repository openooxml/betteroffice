# Agent evaluations

Evaluations of the `@betteroffice/agents` document tools as a model sees them:
through the MCP server, with no other way to read or write the documents.

- [`formats/`](formats): a scored suite of 10 DOCX, 10 XLSX and 10 PPTX editing
  tasks with automatic checkers.
- [`usability/`](usability): manual DOCX, XLSX and PPTX tasks for reviewing
  agent ergonomics on an installed package.

## Format suite

### What it measures

Whether a model completes a realistic document, spreadsheet or presentation
edit using only the BetterOffice MCP tools, and whether the saved file is
exactly right: the requested change is present, everything else in the
document is preserved, and the input file is unchanged. It measures the tools
and the model together; it is not a general model benchmark.

### Tasks

[`formats/tasks.json`](formats/tasks.json) defines each task: an id, the input
and output file names, the prompt, and an `expected` specification that the
checker applies to the input to derive the correct result.

| Format | Tasks |
| --- | --- |
| DOCX | fix a misspelled heading, replace a term case-sensitively in the title, body, list and table, correct table values, add a table row above the total and update the total, delete a table row and update the total, make a paragraph a Heading 2, bold a phrase inside a paragraph, insert a paragraph below a heading, delete a draft note, a combined placeholder replacement, line deletion and heading change |
| XLSX | correct a value and recalculate totals, fix a total formula, case-sensitive replace across two sheets, style a header, sort rows by a column, add a total row with formulas and formatting, insert a computed column, rename and reorder sheets while keeping cross-sheet formulas working, build a summary from another sheet, a combined column insert, sort and format |
| PPTX | change a title, update a figure in a bullet, replace a bullet list, replace a term in titles and bullets but not notes, change two titles but not matching bullet text, delete a slide, reorder slides with their notes, replace speaker notes, add a positioned text box, insert a slide with a title and bullets |

Prompts state the goal, not the tool sequence, and name details that must be
preserved (other figures, notes, formatting or order). The checker compares the
whole document, so an edit that looks right but damages the rest of the file
fails.

### Setup

[`formats/run.py`](formats/run.py) runs every task in a fresh temporary
directory that contains only the input file. It starts `codex exec` with:

- the model under test (`-m`);
- a read-only sandbox, with the shell, browser, web search, image, plugin
  and every other codex tool that could read local files disabled (the full
  list is in `run.py`);
- the code-mode host left enabled, because models that call tools only through
  code mode have no tools without it; it is a V8 isolate without file or network
  access, and each tool call it makes is logged and checked like a direct call;
- the BetterOffice MCP server as the only configured tool server, rooted at the
  task directory; the server runs outside the sandbox and is the only process
  that can write the output. It is marked `required` with a 60 second
  `startup_timeout_sec`: codex otherwise waits only
  `mcp_optional_startup_grace_ms` (1 second) for a server before the first
  turn and starts without its tools, while a required server is waited for and
  a server that cannot start ends the agent with an error;
- `--ignore-user-config`, `--ignore-rules` and `--ephemeral`, to skip user
  configuration, execution rules and persisted session state;
- no approval prompts and a 600 second timeout per task.

The prompt requires BetterOffice MCP tools for all document access. Before the
first task the runner starts the server once with an empty root, completes the
MCP handshake and `tools/list`, and stops with an error unless every tool the
tasks need is listed within 60 seconds; `results.json` records the measured
startup time as `mcp_startup_seconds`.

### Generic arm

`--arm generic --python <python>` runs the same tasks, prompts and checker with
the same model but without the BetterOffice MCP server: the agent reads and
edits the documents with shell commands and Python with openpyxl, python-pptx
and python-docx. The runner first checks that `<python>` imports the three
libraries and puts its directory first on `PATH`. codex runs with
`--sandbox workspace-write` in the task directory, network access off, no
additional writable roots and no MCP servers. `shell_tool`, `unified_exec` and
`unified_exec_tty` are enabled; every other feature listed in `run.py` stays
disabled, including `shell_snapshot`, so commands inherit the runner's
environment. Shell commands and file changes are allowed and counted as tool
calls; an MCP call, a web search or any other tool item fails the task.
Each task starts in a fresh temporary directory outside the repository containing
only its input file; the generic arm is told to use only this directory, the Python
interpreter and its standard library/site-packages. The audit fails logged commands
that reference the repository, evaluation sources or task metadata, and file changes
outside the task directory. The workspace-write sandbox restricts writes and disables
network access. Reads are not sandboxed: this is an audit of logged commands, not
filesystem isolation.

### Scoring

A task passes only if all of these hold:

- the agent exits with code 0 within the timeout;
- the agent makes at least one tool call;
- the JSON event log contains no `command_execution`, MCP call to a server
  other than `betteroffice`, or other tool-call item; only BetterOffice MCP
  items, agent messages, reasoning and errors are allowed (the generic arm
  allows shell commands and file changes instead of MCP items), and an unknown
  item type fails the task with its reason in both reports;
- the output file exists and is a valid package: intact ZIP, well-formed XML,
  and every relationship target and content type override present;
- [`formats/check.py`](formats/check.py) finds the expected document: it
  derives the expected model from the input and the task's `expected` block and
  compares it with the output. For XLSX that covers sheet names, order and
  visibility, cell values and cached results, formulas, number formats, fonts,
  fills, borders, alignment, protection and row properties, plus every workbook
  attribute and child, including all view attributes, defined names, workbook
  properties and calculation properties. `activeTab` and `firstSheet`, including
  their default index 0, and local defined-name scopes follow their sheets
  through renames and reorders;
  defined-name references receive the expected renames. For PPTX it covers
  slide size and order, layouts, masters and themes, backgrounds, every
  attribute and child of each existing shape (geometry, properties, style and
  text body), paragraph, run, field and end paragraph properties, bullets, text
  and speaker notes, plus every presentation attribute and child, presentation
  properties, view properties, notes masters and table styles. Every run
  counts, including an empty run and its formatting, and only paragraphs whose
  text the task changes differ from the input. For DOCX it covers the body in
  order: each paragraph's style (no style means the default paragraph style),
  paragraph properties including numbering, every run's properties and text,
  including `xml:space` for leading or trailing whitespace
  (adjacent runs with identical properties merged, empty runs with formatting
  kept) and every other paragraph child; each table's properties, grid, rows
  and cells with their properties and content; the section properties and the
  document element. Styles, numbering, settings and headers or footers are
  compared as package parts. Tracked-change markup the input does not have
  (`w:ins`, `w:del`, moves, property changes) fails the task. A new paragraph
  is expected in the default paragraph style without direct formatting, a new
  table row copies the row it is modeled on, including row, cell, paragraph and
  run properties, and bold sets `w:b` on exactly the requested phrase. All
  formats compare all remaining package parts, including document properties
  and relationships;
- the input file is byte-identical to the original.

The score is the number of passing tasks out of 30. The runner also records
the duration and the number of tool calls per task (BetterOffice MCP calls, or
shell commands and file changes in the generic arm), and `results.md` counts
the tasks that made no tool call.

ZIP metadata, XML serialization, relationship IDs, style and string table
indexes, worksheet `dimension` and row `spans` are normalized because saves
can regenerate those representations without changing document state.
Tab selection is compared per workbook view as the set of selected sheet
names plus the view's active tab, because an absent `tabSelected` equals `0`
and readers always treat the active tab as selected.
Cell, row and column formats and named cell styles are compared resolved
through the number format, font, fill, border, cell style format and cell
format pools, never by pool index, so those pools may be renumbered or
reordered. Named cell styles (`cellStyles`) are matched by name. Differential
formats (`dxfs`) are compared as written: no fixture has a conditional format
or table style that references them, and the BetterOffice writer keeps that
pool as read.
An `apply*` flag (absent means applied) counts only where the format differs
from its cell style, and a fill's `bgColor` only for patterned fills, because
readers ignore them elsewhere.
`xl/calcChain.xml` and its relationship and content type are excluded because
formula edits can invalidate, rebuild or remove that calculation cache.
A referenced notes slide is compared with its slide by content type,
relationships (including the link back to that slide), shapes, text, and
paragraph, run, field and end paragraph properties, but not by part name.
Replaced notes must keep the paragraph, run and end paragraph properties of
the source notes paragraphs, which must be uniform, and a blank line keeps one
empty run with those properties; the formatting of created notes is not
predicted. Unreferenced slide and note payloads are excluded because a slide
deletion can retain or remove those detached parts without changing the
presentation.
In DOCX, the value of `dcterms:modified` in `docProps/core.xml` is ignored
because the BetterOffice writer stamps the save time there; the element itself
is compared. Otherwise only equivalent representations are equal: an empty
property element equals an absent one, adjacent runs with identical properties
are merged, and a paragraph without a style has the default paragraph style.
The DOCX fixtures carry `w14:paraId` and `w14:textId` on every paragraph, as Word
writes them. Paragraphs from the input keep both exactly, edited or not. A new
paragraph (an inserted paragraph or the paragraphs of an inserted row) may
carry fresh ids or none. Duplicate paraIds or duplicate textIds fail. Rsid attributes are compared as
written; the fixtures have none.

The checkers are validated against known answers:
[`formats/validate-checkers.py`](formats/validate-checkers.py) builds a correct
output for every task and checks that it passes, and that the untouched input,
an output with an unrelated edit, and an output with a damaged requested change
all fail, plus unrelated workbook-, presentation- and document-level edits;
two enriched fixtures also check view and defined-name remapping and metadata
preservation, and targeted variants check tab selection, notes parts, slide
shape elements, paragraph end properties, fields and empty formatted runs,
cell, row, column and named style formats, and for DOCX tracked-change markup,
run splits and formatting, row placement, style resolution, reordered content,
package parts and paragraph ids (65 accepted, 30 untouched rejected, 157
corrupted rejected).
[`formats/test-runner.py`](formats/test-runner.py) tests the runner itself with
a stand-in agent.

### Fixtures

The 30 input files are synthetic documents generated by
[`formats/generate-fixtures.ts`](formats/generate-fixtures.ts) with fixed
timestamps; the DOCX files are hand-written OOXML opened and saved by the
BetterOffice engine, with paragraph ids on every paragraph. They are not
stored in the repository. They are published in the public corpus bucket under
`https://corpus.betteroffice.dev/eval/formats/v2/`, and
[`formats/fixtures.json`](formats/fixtures.json) pins each file's size and
SHA-256, and its `provenance` note states this origin and how to regenerate.

The checkers download a fixture on first use, verify it, and keep it in the
asset cache (`QUALITY_ASSET_CACHE`, default
`~/.cache/betteroffice/eval-assets`). Set `FORMAT_EVAL_FIXTURES=<dir>` to read
them from a local directory instead; they are verified the same way.

`generate-fixtures.ts` prints each file's size and SHA-256. Regenerated files
can differ in bytes when engine output changes, so a changed fixture set is
published under a new version prefix and `fixtures.json` is updated with it.

### Running

Build the packages the MCP server needs, then run the suite with any model
that `codex exec -m` accepts:

```sh
for p in fonts docx xlsx pptx agents; do bun run --filter "@betteroffice/$p" build; done
eval/formats/run.sh "node packages/agents/dist/cli.js" /tmp/format-eval gpt-6-luna
```

The generic arm takes no MCP command:

```sh
eval/formats/run.sh --arm generic --python /path/to/venv/bin/python /tmp/format-eval-generic gpt-6-luna
```

Either arm takes `--formats xlsx,pptx` (comma-separated, default all) to run
only those formats' tasks; the pre-flight then requires only their tools, and
both reports name the formats that ran.

The results directory must be new or empty and outside the repository. It
receives `results.md`, `results.json` (with the arm), and per task the agent's JSON event log,
its stderr and the output file. Check a single output or the checkers alone
with:

```sh
python3 eval/formats/check.py xlsx-04-style-header /tmp/format-eval/xlsx-04-style-header/output.xlsx
python3 eval/formats/validate-checkers.py
python3 eval/formats/test-runner.py
bun eval/formats/generate-fixtures.ts /tmp/format-fixtures
```

## Usability tasks

These are a small sample for reviewing how an agent works with the tools, not a
benchmark. Run each task in a fresh agent session with only the installed
BetterOffice MCP tools for document access, with the model selected
explicitly (for example `codex exec -m gpt-6-luna`). Keep transcripts and
screenshots outside git.

Build and pack `@betteroffice/docx`, `@betteroffice/fonts`, `@betteroffice/xlsx`,
`@betteroffice/pptx` and `@betteroffice/agents`, and install them into a
separate directory. When testing unreleased builds, use package-manager
overrides so the packed engines are used throughout the install; otherwise a
nested published engine may lack the new APIs. Point the MCP client at the
installed `betteroffice-mcp` with `--root <task directory>` and verify
`tools/list` before starting.

### DOCX

```sh
bun eval/usability/prepare.ts /tmp/office-eval
```

The three documents contain 1,506 paragraphs across the body, a table cell and
a header. Give each session one task:

1. **Revenue:** In revenue.docx, update executive-summary revenue from
   €4.2 million to €5.1 million and delivery from 15 October 2026 to
   12 November 2026. Preserve formatting and all other content. Inspect the
   proposed first page, verify reopening, and export revenue-updated.docx.
2. **Repeated text and table:** In wording.docx, find “The forecast is draft;
   the appendix is draft.” Replace only the second “draft” with “approved”.
   Change the cell “Table target: approved” to “Table target: pending review”.
   Inspect, verify reopening, and export wording-updated.docx.
3. **Tracked changes across stories:** In cross-story.docx, replace the leading
   word “Internal” of the header with “Board review”, keeping the rest of the
   header, and the heading “Executive summary” with “Board summary”. Inspect a proposed page,
   then accept as native tracked changes attributed to “Luna Review” at
   `2026-09-27T12:00:00Z`. Export cross-story-updated.docx and verify reopening.

Require the source to remain untouched, and ask the agent to report failed
tool calls, confusing responses and unverified claims. Do not tell it how to
construct offsets or which tool sequence to use. Use a fresh directory for
repeats, because exports never overwrite files. Validate the files
independently of the agent's account:

```sh
python3 eval/usability/check.py /tmp/office-eval
```

### XLSX and PPTX

```sh
bun eval/usability/prepare-formats.ts /tmp/betteroffice-formats-eval
```

Run each task with the MCP CLI rooted at its directory:

- `xlsx`: set Budget!B3 to 1000 and E3 to `=D3*2`; review, verify, export
  `budget-revised.xlsx`, and inspect its calculated result.
- `pptx`: change revenue to €5.1 million and only the second occurrence in
  “Risk risk.” to “opportunity”; review, verify, export `slides-revised.pptx`.
- `mixed`: set Budget!B3 to 700 and presentation revenue to €6.0 million;
  review, verify, export `budget-revised.xlsx` and `slides-revised.pptx`.

Require original files, formatting and unrelated content to remain intact, then
validate the outputs:

```sh
python3 eval/usability/check-formats.py /tmp/betteroffice-formats-eval
```

If a run fails because the server was unavailable, fix installation or
discovery and restart; it does not count as a task result.
