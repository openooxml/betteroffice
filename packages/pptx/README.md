# @betteroffice/pptx

Framework-free core for the BetterOffice PPTX editor — the Rust parser, yrs deck
model, slide layout, and display-list engine compiled to WebAssembly, plus the
Canvas2D replay host.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/pptx
```

Most apps want the turnkey React component in
[`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react).
Use this package directly to build custom presentation chrome.

## Usage

```ts
import { initWasm, openPresentation, paintSlide } from '@betteroffice/pptx';

await initWasm();
const bytes = new Uint8Array(await file.arrayBuffer());
const deck = openPresentation(bytes, {
  fonts: [{ family: 'My Sans', bytes: fontBytes }],
});
const frame = deck.layoutSlide(0);
await paintSlide(canvas.getContext('2d')!, frame, devicePixelRatio);
```

All parsing, edits, collaboration state, text shaping, layout, hit-testing, and
display-list emission stay in Rust. The package decodes the typed boundary and
replays the resulting primitives on canvas. Font bytes are supplied by the host
and registered with the Rust shaper through `openPresentation`.

## Collaboration

`PresentationHandle` is a collaboration replica. Pair it with
`CollaborationProvider` and a transport implementing the small
`CollaborationTransport` interface. The provider speaks the standard Yjs sync-v1
wire protocol, performs state-vector handshakes, forwards only local updates,
and bounds frames and pending backpressure bytes.

```ts
import { CollaborationProvider } from '@betteroffice/pptx';

const provider = new CollaborationProvider(deck, transport);
deck.onUpdate((_update, origin) => {
  if (origin === 'remote') repaint();
});
provider.connect();
```

## Development

The generated `.wasm` binary is intentionally not committed. From the repository
root, install `wasm-pack` 0.15.0 and run `bun scripts/build-pptx-wasm.ts`.
Package builds copy the binary into `dist/generated`.

Docs: https://betteroffice.dev · Apache-2.0.
