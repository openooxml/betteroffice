<p align="center">
  <a href="https://betteroffice.dev">
    <img src="./.github/assets/header.svg" alt="BetterOffice: the open-source office suite, built on native OOXML engines in Rust" width="100%">
  </a>
</p>

<p align="center">
  Documents, spreadsheets, and presentations — engines in Rust, editors in TypeScript, everything client-side.
</p>

<p align="center">
  <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-4ade80.svg?style=flat-square" alt="license"></a>
  <a href="https://www.npmjs.com/org/betteroffice"><img src="https://img.shields.io/endpoint?url=https%3A%2F%2Fbetteroffice.dev%2Fapi%2Fnpm-downloads&amp;style=flat-square&amp;logo=npm&amp;label=downloads&amp;color=CB3837&amp;cacheSeconds=86400" alt="npm downloads"></a>
  <a href="https://betteroffice.dev"><img src="https://img.shields.io/badge/betteroffice.dev-0a0a0a?style=flat-square" alt="betteroffice.dev"></a>
  <a href="https://openooxml.org"><img src="https://img.shields.io/badge/openooxml.org-0a0a0a?style=flat-square" alt="openooxml.org"></a>
</p>

## Engines

The document engines are Rust crates compiled to WebAssembly for the browser and running natively on servers. Every value from a user file is treated as attacker-controlled; guards are enforced by construction at the trust boundaries.

Rust consumers should use `betteroffice-xlsx`; the lower-level `xlsx-*` crates are its engine components.

| crate | what it does |
|---|---|
| [`ooxml-opc`](crates/ooxml-opc) | OPC (zip) container read/write with decompression-bomb and path-traversal guards — shared by every format |
| [`ooxml-text`](crates/ooxml-text) | shared font storage, shaping, bidi, line breaking, and glyph outlines |
| [`docx-layout`](crates/docx-layout) | DOCX pagination and display-list generation |
| [`docx-wasm`](crates/docx-wasm) | the wasm boundary for the document engine |
| [`pptx-render`](crates/pptx-render) | PPTX composed-slide to display-list compiler |
| [`pptx-wasm`](crates/pptx-wasm) | the wasm boundary for presentation rendering |
| [`betteroffice-xlsx`](crates/betteroffice-xlsx) | typed XLSX editor, calculation, rendering, and save facade for Rust |
| [`xlsx-model`](crates/xlsx-model) | workbook, cells, addresses, dates, styles, number formats |
| [`xlsx-parse`](crates/xlsx-parse) | streaming SpreadsheetML parse and serialize |
| [`xlsx-calc`](crates/xlsx-calc) | formula engine: parser, dependency graph, incremental recalc |
| [`xlsx-ops`](crates/xlsx-ops) | invertible edit operations, undo, address remapping, proposals |
| [`xlsx-render`](crates/xlsx-render) | grid geometry and display list |
| [`xlsx-raster`](crates/xlsx-raster) | headless raster backend (tiny-skia), pixel-identical everywhere |
| [`xlsx-wasm`](crates/xlsx-wasm) | the wasm boundary for the spreadsheet engine |
| [`xlsx-cli`](crates/xlsx-cli) | render and inspect workbooks from the command line |

## Packages

The editor packages will ship under the `@betteroffice` scope: a framework-free core per format, with thin framework adapters on top.

| package | |
|---|---|
| `@betteroffice/docx` | word processing core |
| `@betteroffice/docx-react` | React chrome for the docx editor |
| `@betteroffice/xlsx` | spreadsheet core |
| `@betteroffice/xlsx-react` | React chrome for the spreadsheet |
| `@betteroffice/pptx` | presentation core |
| `@betteroffice/pptx-react` | React chrome for presentations |

## Structure

- `crates/` — the Rust engines
- `packages/` — the TypeScript editor packages
- `apps/web` — [betteroffice.dev](https://betteroffice.dev) (Next.js on Cloudflare Workers)
- `apps/docs` — documentation

## Development

```bash
bun install
bun run build:xlsx-wasm # compile the ignored spreadsheet wasm asset
bun run dev          # web app
bun run rust:check   # fmt + clippy + tests for the engines
```

## Contributing

Contributions are welcome. We ask for a one-time signature of the [Contributor License Agreement](CLA.md) on your first pull request ([corporate version](CCLA.md)).

## License

[Apache-2.0](LICENSE) — third-party attribution in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
