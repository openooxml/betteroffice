import { beforeAll } from 'bun:test';
import { readFileSync } from 'node:fs';
import { describeDisposeTable } from '../../../../bindings/dispose-cases';
import * as engine from './engine';

const deck = new Uint8Array(readFileSync(new URL('../../../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)));

beforeAll(async () => {
  await engine.default({ module_or_path: readFileSync(new URL('./generated/pptx_wasm_bg.wasm', import.meta.url)) });
}, 30_000);

describeDisposeTable(engine.engineKind, engine, [
  { name: 'PptxDocument', engineClass: engine.PptxDocument, make: () => engine.PptxDocument.openCollaborative(deck, 8), use: (document) => (document as engine.PptxDocument).canUndo() },
  { name: 'PptxRenderer', engineClass: engine.PptxRenderer, make: () => new engine.PptxRenderer(), use: (renderer) => (renderer as engine.PptxRenderer).hitTestJson(0, 0) },
]);
