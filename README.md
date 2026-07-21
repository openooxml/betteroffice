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
  <a href="https://betteroffice.dev"><img src="https://betteroffice.dev/api/badge?label=&amp;message=betteroffice.dev&amp;color=0a0a0a" alt="betteroffice.dev"></a>
  <a href="https://openooxml.org"><img src="https://betteroffice.dev/api/badge?label=&amp;message=openooxml.org&amp;color=0a0a0a" alt="openooxml.org"></a>
</p>

## Packages

| package | what it does |
|---|---|
| [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) | framework-free .docx editor core — parsing, CRDT editing, and page layout in Rust through WebAssembly |
| [`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react) | drop-in React .docx editor |
| [`betteroffice-xlsx`](https://crates.io/crates/betteroffice-xlsx) | typed Rust API for opening, editing, calculating, rendering, and saving XLSX workbooks |
| [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) | framework-free spreadsheet core powered by the Rust engine through WebAssembly |
| [`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react) | drop-in React spreadsheet editor |
| [`betteroffice-pptx`](https://crates.io/crates/betteroffice-pptx) | typed Rust API for opening, editing, rendering, and saving PPTX presentations |
| [`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) | framework-free .pptx editor core — slide model, masters, and rendering in Rust through WebAssembly |
| [`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react) | drop-in React .pptx editor |

## Structure

- `crates/` — the Rust engines
- `packages/` — the TypeScript editor packages
- `apps/web` — [betteroffice.dev](https://betteroffice.dev) (Next.js on Cloudflare Workers)
- `apps/docs` — documentation

## Development

```bash
bun install
bun run build:xlsx-wasm # compile the ignored spreadsheet wasm asset
bun run build:docx-wasm # compile the ignored document wasm assets
bun run dev          # web app
bun run rust:check   # fmt + clippy + tests for the engines
```

## Contributing

Contributions are welcome. We ask for a one-time signature of the [Contributor License Agreement](CLA.md) on your first pull request ([corporate version](CCLA.md)).

## License

[Apache-2.0](LICENSE) — third-party attribution in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
