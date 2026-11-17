import { beforeAll } from 'bun:test';
import { readFileSync } from 'node:fs';
import { describeDisposeTable } from '../../../../bindings/dispose-cases';
import * as engine from './engine';

const workbook = new Uint8Array(readFileSync(new URL('../../../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url)));

beforeAll(async () => {
  await engine.default({ module_or_path: readFileSync(new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url)) });
}, 30_000);

describeDisposeTable(engine.engineKind, engine, [
  { name: 'XlsxDocument', engineClass: engine.XlsxDocument, make: () => engine.XlsxDocument.open(workbook), use: (document) => (document as engine.XlsxDocument).documentVersion() },
  { name: 'XlsxSnapshotBuilder', engineClass: engine.XlsxSnapshotBuilder, make: () => new engine.XlsxSnapshotBuilder(), use: (builder) => (builder as engine.XlsxSnapshotBuilder).advance(1, 2 ** 30) },
]);
