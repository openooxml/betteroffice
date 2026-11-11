import { beforeAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describeHeapTable } from '../../../../bindings/heap-cases';
import { preloadEditWasm } from './edit';
import * as engine from './engineEdit';
import { wasmModuleMemories } from './loadWasmAsset';

let docx: Uint8Array;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url))));
  const encoder = new TextEncoder();
  const paragraphs = Array.from({ length: 400 }, (_, index) =>
    `<w:p><w:r><w:t>Paragraph ${index} with enough text to hold some heap in the editing core.</w:t></w:r></w:p>`).join('');
  docx = engine.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`),
  });
}, 30_000);

function openSession(): engine.EditSession {
  const session = new engine.EditSession(31);
  try {
    session.open_docx(docx, true, 'opening');
  } catch (error) {
    session.free();
    throw error;
  }
  return session;
}

describeHeapTable(
  engine.engineKind,
  () => engine,
  openSession,
  (session) => session.free(),
  () => wasmModuleMemories().find(({ label }) => label === 'docx-edit'),
);

test('startHeapCounting from the package entry starts counting in a fresh process', () => {
  const script = [
    `const edit = await import(${JSON.stringify(fileURLToPath(new URL('./edit.ts', import.meta.url)))});`,
    `const engine = await import(${JSON.stringify(fileURLToPath(new URL('./engineEdit.ts', import.meta.url)))});`,
    `const yrs = await import(${JSON.stringify(fileURLToPath(new URL('../yrs/index.ts', import.meta.url)))});`,
    `const { readFileSync } = await import('node:fs');`,
    `await edit.preloadEditWasm(new Uint8Array(readFileSync(${JSON.stringify(fileURLToPath(new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url)))})));`,
    'const before = engine.wasm_heap_counted();',
    'const started = yrs.startHeapCounting();',
    'const session = new engine.EditSession(41);',
    'const live = engine.wasm_live_bytes();',
    'session.free();',
    'console.log(JSON.stringify({ kind: engine.engineKind, before, started, after: engine.wasm_heap_counted(), live }));',
  ].join('\n');
  const result = spawnSync(process.execPath, ['--eval', script], { cwd: import.meta.dir, encoding: 'utf8', timeout: 120_000 });
  expect(result.status).toBe(0);
  const report = JSON.parse(result.stdout.trim().split('\n').at(-1) as string);
  expect(report.kind).toBe(engine.engineKind);
  expect(report.before).toBe(engine.engineKind === 'wasm');
  expect(report.started).toBe(true);
  expect(report.after).toBe(true);
  expect(report.live).toBeGreaterThan(0);
});
