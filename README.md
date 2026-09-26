<p align="center">
  <a href="https://betteroffice.dev">
    <img src="./.github/assets/header.svg" alt="BetterOffice: the open-source office suite, built on native OOXML engines in Rust" width="100%">
  </a>
</p>

<p align="center">
  Rust-native OOXML engines with collaboration and agent editing at the core.<br>
  WebAssembly for browsers. Headless APIs for servers. Native Rust where you need it.
</p>

<p align="center">
  <a href="./LICENSE"><img src="https://betteroffice.dev/api/badge?label=license&amp;message=Apache-2.0&amp;color=4ade80" alt="license"></a>
  <a href="https://www.npmjs.com/org/betteroffice"><img src="https://betteroffice.dev/api/npm-downloads-badge" alt="npm downloads"></a>
  <a href="https://crates.io/search?q=betteroffice"><img src="https://betteroffice.dev/api/crates-downloads-badge" alt="crates.io downloads"></a>
  <a href="https://pypi.org/project/betteroffice-xlsx/"><img src="https://betteroffice.dev/api/pypi-downloads-badge" alt="PyPI downloads"></a>
  <a href="https://betteroffice.dev"><img src="https://betteroffice.dev/api/badge?label=&amp;message=betteroffice.dev&amp;color=0a0a0a" alt="betteroffice.dev"></a>
</p>

## Features

- **Documents, spreadsheets, and slides.** Open, edit, render, and save DOCX,
  XLSX and PPTX files with high fidelity. Native OOXML editing preserves untouched
  file parts losslessly when round-tripping.

- **Agent editing with human review.** Review attributed agent edits through
  tracked changes, inline diffs, and before-and-after previews. Accept or
  reject changes directly in the editor.

- **Real-time collaboration.** People and agents edit the same file together,
  with live cursors and selections. Concurrent changes merge automatically,
  and offline edits sync when peers reconnect.

- **Undo and redo.** Navigate editing history and undo accepted agent
  proposals as a single step.
  DOCX hosts can place explicit undo boundaries and choose automatic typing
  coalescence or manual grouping across multiple edits.

- **Embed or automate.** Drop React editors into your app, build on the
  framework-free JavaScript cores, or use Rust and Python APIs for headless
  processing and agent workflows.

- **Host-controlled PPTX editing.** PPTX hosts can intercept saving, flush accepted input, query slide content
  under the pointer, group undo with explicit boundaries, and reposition comments.

- **Open source and self-hostable.** Apache-2.0 licensed, with control over
  your document storage, deployment, and collaboration infrastructure.

- **Host-controlled DOCX saving.** Intercept Save before export and await pending
  editor input before reading, changing, or persisting the document.

- **Version-checked DOCX edit batches.** DOCX JavaScript and React hosts read
  paragraphs with a session version, then apply atomic text and paragraph
  batches that commit as one undo step or return a typed refusal with nothing
  changed. Text steps stay within one paragraph and paragraph steps need a
  document opened from DOCX bytes.

- **Version-checked XLSX edit batches.** XLSX hosts in JavaScript, React, Rust
  and Python read cells with a session version, then apply atomic cell input,
  formula, number-format and style batches that commit as one recalculated undo
  step or return a typed refusal with nothing changed. Batches do not insert or
  delete rows, columns or sheets, and refuse writes to merged-cell followers,
  array-formula cells and protected sheets.

- **Structured XLSX export.** XLSX exports bounded sparse worksheet content and
  Markdown with positional anchors, formulas, stored values, formatted text,
  explicit hidden-content options, and omission diagnostics. Export does not
  recalculate formulas.

[Try it out](https://demo.betteroffice.dev), or
[explore the docs](https://docs.betteroffice.dev) for setup, APIs, and format support.

## Packages

### Documents — `.docx`

| package | registry | what it does |
|---|---|---|
| [`betteroffice-docx`](https://crates.io/crates/betteroffice-docx) | crates.io | typed Rust API for opening, editing, laying out, and saving DOCX documents |
| [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) | npm | framework-free .docx editor core — parsing, CRDT editing, and page layout in Rust through WebAssembly |
| [`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react) | npm | drop-in React .docx editor |
| [`betteroffice-docx`](https://pypi.org/project/betteroffice-docx/) | PyPI | Python API for reading, editing, laying out, and rasterizing DOCX documents |

### Spreadsheets — `.xlsx`

| package | registry | what it does |
|---|---|---|
| [`betteroffice-xlsx`](https://crates.io/crates/betteroffice-xlsx) | crates.io | typed Rust API for opening, editing, calculating, rendering, and saving XLSX workbooks |
| [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) | npm | framework-free spreadsheet core powered by the Rust engine through WebAssembly |
| [`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react) | npm | drop-in React spreadsheet editor |
| [`betteroffice-xlsx`](https://pypi.org/project/betteroffice-xlsx/) | PyPI | Python API for opening, recalculating, styling, rendering, and saving XLSX workbooks |

### Presentations — `.pptx`

| package | registry | what it does |
|---|---|---|
| [`betteroffice-pptx`](https://crates.io/crates/betteroffice-pptx) | crates.io | typed Rust API for opening, editing, rendering, and saving PPTX presentations |
| [`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) | npm | framework-free .pptx editor core — slide model, masters, image insertion, shape ordering, and rendering in Rust through WebAssembly |
| [`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react) | npm | drop-in React .pptx editor |
| [`betteroffice-pptx`](https://pypi.org/project/betteroffice-pptx/) | PyPI | Python API for reading, editing, saving, and laying out PPTX presentations |

### Diagrams (`.vsdx`, source preview)

VSDX support is available from source and is not yet published.

What to install for which language, with a first example each:
[npm](https://docs.betteroffice.dev/docs/javascript),
[crates.io](https://docs.betteroffice.dev/docs/rust),
[PyPI](https://docs.betteroffice.dev/docs/python).

<!-- BEGIN GENERATED VISUAL FIDELITY -->
## Benchmarks

*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml).*

### DOCX

Page agreement is reported on its own because a document either paginates as Word does or it does not; SSIM cannot express that.

<table>
<tr><th width="180"></th><th width="230" align="right">BetterOffice (<a href="https://www.npmjs.com/package/@betteroffice/docx/v/0.2.1">0.2.1</a>)</th><th width="230" align="right">BetterOffice (<a href="https://github.com/openooxml/betteroffice/commit/e8ce123fa9f8e9efc5096811fb98f6f0c17d2e48">e8ce123f</a>)</th><th width="230" align="right">LibreOffice (26.2.3.2)</th></tr>
<tr><td>Exact page counts</td><td align="right">48/63</td><td align="right">63/63</td><td align="right">45/63</td></tr>
<tr><td>Absolute page error</td><td align="right">32</td><td align="right">0</td><td align="right">56</td></tr>
<tr><td>SSIM</td><td align="right">0.7540</td><td align="right">0.8368</td><td align="right">0.7749</td></tr>
<tr><td>Scored/total</td><td align="right">63/63</td><td align="right">63/63</td><td align="right">63/63</td></tr>
<tr><td>Render time (avg)</td><td align="right">546 ms</td><td align="right">527 ms</td><td align="right">788 ms</td></tr>
<tr><td>Parse success</td><td align="right">100.00%</td><td align="right">100.00%</td><td align="right">100.00%</td></tr>
</table>


### PPTX

<table>
<tr><th width="180"></th><th width="230" align="right">BetterOffice (<a href="https://www.npmjs.com/package/@betteroffice/pptx/v/0.1.1">0.1.1</a>)</th><th width="230" align="right">BetterOffice (<a href="https://github.com/openooxml/betteroffice/commit/e8ce123fa9f8e9efc5096811fb98f6f0c17d2e48">e8ce123f</a>)</th><th width="230" align="right">LibreOffice (26.2.3.2)</th></tr>
<tr><td>SSIM</td><td align="right">0.8667</td><td align="right">0.8924</td><td align="right">0.9043</td></tr>
<tr><td>Scored/total</td><td align="right">88/103</td><td align="right">103/103</td><td align="right">103/103</td></tr>
<tr><td>Render time (avg)</td><td align="right">205 ms</td><td align="right">144 ms</td><td align="right">1310 ms</td></tr>
<tr><td>Parse success</td><td align="right">100.00%</td><td align="right">100.00%</td><td align="right">100.00%</td></tr>
</table>

### XLSX

<table>
<tr><th width="180"></th><th width="230" align="right">BetterOffice (<a href="https://www.npmjs.com/package/@betteroffice/xlsx/v/0.2.1">0.2.1</a>)</th><th width="230" align="right">BetterOffice (<a href="https://github.com/openooxml/betteroffice/commit/e8ce123fa9f8e9efc5096811fb98f6f0c17d2e48">e8ce123f</a>)</th><th width="230" align="right">LibreOffice (26.2.3.2)</th></tr>
<tr><td>SSIM</td><td align="right">0.7002</td><td align="right">0.7660</td><td align="right">0.6901</td></tr>
<tr><td>Scored/total</td><td align="right">205/208</td><td align="right">207/208</td><td align="right">204/208</td></tr>
<tr><td>Recalc accuracy</td><td align="right">68.28%</td><td align="right">99.60%</td><td align="right">85.87%</td></tr>
<tr><td>Recalc time (avg)</td><td align="right">70 ms</td><td align="right">89 ms</td><td align="right">466 ms</td></tr>
<tr><td>Parse success</td><td align="right">99.52%</td><td align="right">100.00%</td><td align="right">100.00%</td></tr>
</table>

For scoring, timing, coverage, and limitations, see the [benchmark methodology](scripts/office-quality/README.md).

<!-- END GENERATED VISUAL FIDELITY -->

## Structure

- `crates/` — the Rust engines
- `packages/` — the TypeScript editor packages
- `bindings/` — the Python bindings
- `e2e/` — browser tests, corpus scenarios, and the shared test harness
- `apps/web` — [betteroffice.dev](https://betteroffice.dev) (Next.js on Cloudflare Workers)
- `apps/demo` — editor playground
- `apps/docs` — documentation

## Development

```bash
bun install
bun run dev:demo     # editor playground; builds the wasm bundles on first run
bun run dev          # betteroffice.dev site (no wasm needed)
bun run dev:docs     # documentation site
bun run rust:check   # fmt + clippy + tests for the engines
```

Run `bun run test:e2e:browser` for headless browser tests of all three editors.
The [E2E guide](e2e/README.md) also covers pinned corpus scenarios,
cross-SDK checks, and operation profiles. The XLSX and PPTX JavaScript cores expose
opt-in operation timings; the native XLSX facade also reports edit stage timings.

Use the [Office visual quality harness](scripts/office-quality/README.md) to export Word, PowerPoint, and Excel references, compare local renders, or refresh the fidelity scores with the manual action.



## Contributing

Contributions are welcome. We ask for a one-time signature of the [Contributor License Agreement](CLA.md) on your first pull request ([corporate version](CCLA.md)).

## License

[Apache-2.0](LICENSE) — third-party attribution in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
