# Agent usability evaluation

Run these tasks with a fresh agent context and only the installed BetterOffice
MCP tools for document access. The fixtures contain 1,506 paragraphs across the
body, a table cell, and a header. Keep run transcripts and rendered screenshots
outside git; attach review evidence to the PR.

```sh
bun packages/agents/eval/prepare.ts /tmp/office-eval
```

Build and pack `@betteroffice/docx`, `@betteroffice/fonts`, and
`@betteroffice/agents`, then install them into a separate directory. When testing
unreleased builds, use package-manager overrides to force the packed DOCX/fonts
dependencies throughout the install; otherwise a nested published engine may
lack the new APIs. Point an MCP client at the installed `betteroffice-mcp` with
`--root /tmp/office-eval`. Verify `tools/list` before starting model runs.

Start new Codex sessions with the model selected explicitly, for example
`codex exec -m gpt-6-luna`, and give each session one task:

1. **Revenue:** In revenue.docx, update executive-summary revenue from
   €4.2 million to €5.1 million and delivery from 15 October 2026 to
   12 November 2026. Preserve formatting and all other content. Inspect the
   proposed first page, verify reopening, and export revenue-updated.docx.
2. **Repeated text and table:** In wording.docx, find “The forecast is draft;
   the appendix is draft.” Replace only the second “draft” with “approved”.
   Change the cell “Table target: approved” to “Table target: pending review”.
   Inspect, verify reopening, and export wording-updated.docx.
3. **Tracked changes across stories:** In cross-story.docx, replace the header
   “Internal — Project Aurora” with “Board review — Project Aurora” and the
   heading “Executive summary” with “Board summary”. Inspect a proposed page,
   then accept as native tracked changes attributed to “Luna Review” at
   `2026-09-27T12:00:00Z`. Export cross-story-updated.docx and verify reopening.

For every task, require the source to remain untouched and ask the agent to
report failed tool calls, confusing responses, and unverified claims. Do not
tell it how to construct offsets or which tool call sequence to use. A fresh
directory is needed for repeats because exports never overwrite files.

Validate the files independently of the agent's account:

```sh
python3 packages/agents/eval/check.py /tmp/office-eval
```

Record task completion, failed calls, recovery behavior, expected tool usage,
and validator results in the PR. These are a small usability sample, not a
general model benchmark. If a run fails because the server was unavailable,
fix installation/discovery and restart; do not count it as a document-task pass.

## XLSX/PPTX prototype tasks

Build and pack the agents package with the matching XLSX and PPTX engine builds.
Install those tarballs in an isolated consumer; override transitive engine
versions so the consumer uses the same unreleased engine APIs as the package.

```sh
bun packages/agents/eval/prepare-formats.ts /tmp/betteroffice-formats-eval
```

Run independent Codex `gpt-6-luna` sessions with the installed MCP CLI rooted at
each task directory. Ask for these tasks using only MCP document tools:

- `xlsx`: set Budget!B3 to 1000 and E3 to `=D3*2`; review, verify, export
  `budget-revised.xlsx`, and inspect its calculated result.
- `pptx`: change revenue to €5.1 million and only the second occurrence in
  “Risk risk.” to “opportunity”; review, verify, export `slides-revised.pptx`.
- `mixed`: set Budget!B3 to 700 and presentation revenue to €6.0 million;
  review, verify, export `budget-revised.xlsx` and `slides-revised.pptx`.

Require original files, formatting, and unrelated content to remain intact.
Keep prompts, tool traces, and screenshots outside git. Validate the outputs:

```sh
python3 packages/agents/eval/check-formats.py /tmp/betteroffice-formats-eval
```

These are reproducible smoke tasks, not a general model-quality benchmark.
