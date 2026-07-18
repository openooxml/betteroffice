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
  <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-4ade80.svg?style=flat-square" alt="license"></a>
  <a href="https://www.npmjs.com/org/betteroffice"><img src="https://img.shields.io/endpoint?url=https%3A%2F%2Fbetteroffice.dev%2Fapi%2Fnpm-downloads&amp;style=flat-square&amp;logo=npm&amp;label=downloads&amp;color=CB3837&amp;cacheSeconds=86400" alt="npm downloads"></a>
  <a href="https://crates.io/search?q=betteroffice"><img src="https://img.shields.io/endpoint?url=https%3A%2F%2Fbetteroffice.dev%2Fapi%2Fcrates-downloads%3Fv%3D1&amp;style=flat-square&amp;logo=rust&amp;label=downloads&amp;color=CE412B&amp;cacheSeconds=86400" alt="crates.io downloads"></a>
  <a href="https://betteroffice.dev"><img src="https://img.shields.io/badge/betteroffice.dev-0a0a0a?style=flat-square" alt="betteroffice.dev"></a>
  <a href="https://openooxml.org"><img src="https://img.shields.io/badge/openooxml.org-0a0a0a?style=flat-square" alt="openooxml.org"></a>
</p>

## Packages

| package | what it does |
|---|---|
| [`betteroffice-xlsx`](https://crates.io/crates/betteroffice-xlsx) | typed Rust API for opening, editing, calculating, rendering, and saving XLSX workbooks |
| [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) | framework-free spreadsheet core powered by the Rust engine through WebAssembly |
| [`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react) | drop-in React spreadsheet editor |

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
