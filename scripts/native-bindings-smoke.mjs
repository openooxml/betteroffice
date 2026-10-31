import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const fixtureRoot = process.argv[2]
  ? pathToFileURL(`${resolve(process.argv[2])}/`)
  : new URL('../', import.meta.url);
const fixture = (path) => readFileSync(new URL(path, fixtureRoot));
const docxBytes = fixture('apps/demo/public/betteroffice-demo.docx');
const xlsxBytes = fixture('crates/ooxml-opc/tests/fixtures/sample.xlsx');
const pptxBytes = fixture('apps/demo/public/betteroffice-demo.pptx');

assert.ok(process.env.BETTEROFFICE_EXPORT_KEYS, 'BETTEROFFICE_EXPORT_KEYS names the export names of each wasm entry');
const exportKeys = JSON.parse(readFileSync(process.env.BETTEROFFICE_EXPORT_KEYS, 'utf8'));
for (const [format, entries] of Object.entries(exportKeys)) {
  for (const [entry, keys] of Object.entries(entries)) {
    const name = `@betteroffice/${format}-native${entry.slice(1)}`;
    assert.deepEqual(Object.keys(await import(name)).sort(), keys, name);
  }
}
const { createYrsSession, saveYrsDocx } = await import('@betteroffice/docx-native/yrs');
const { parseDocx, exportDocxMarkdown, fontPreflight } = await import('@betteroffice/docx-native');
const session = await createYrsSession({ clientId: 7 });
const reopened = await createYrsSession({ clientId: 8 });
try {
  session.seedFromDocx(new Uint8Array(docxBytes));
  const first = session.paragraphs('body')[0];
  assert.ok(first);
  session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Native smoke ');
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  reopened.seedFromDocx(saved.bytes);
  assert.deepEqual(reopened.paragraphs('body').map(({ text }) => text), session.paragraphs('body').map(({ text }) => text));
  assert.ok((await parseDocx(saved.bytes)).package.document.content.length > 0);
  assert.ok(await exportDocxMarkdown(saved.bytes, { revisionView: 'accepted' }));
  assert.ok(Array.isArray((await fontPreflight(saved.bytes)).faces));
} finally {
  reopened.destroy();
  session.destroy();
}

const { initWasm: initPptx, openPresentation } = await import('@betteroffice/pptx-native');
await initPptx();
const presentation = openPresentation(pptxBytes);
const snapshot = presentation.snapshot();
assert.ok(snapshot.slides.length > 0);
presentation.registerFont({
  family: 'Arial',
  bytes: fixture('crates/pptx-raster/tests/assets/Carlito-Regular.ttf')
});
const added = presentation.addTextBox(snapshot.slides[0].id, {
  name: 'Native smoke', rect: { x: 100000, y: 200000, width: 3000000, height: 1000000 },
  text: 'Native', style: { fontFamily: 'Arial', fontSizePt: 18 }
});
const shape = presentation.snapshot().slides[0].shapes.find((shape) => shape.id === added.shapeId);
assert.ok(shape);
presentation.insertText(shape.textStories[0].id, 6, ' package API');
assert.ok(presentation.layoutSlide(0).primitives.length > 0);
const reopenedPresentation = openPresentation(presentation.save());
const saved = reopenedPresentation.snapshot();
assert.equal(saved.slides.length, snapshot.slides.length);
const textBox = saved.slides[0].shapes.find((shape) => shape.name === 'Native smoke');
assert.ok(textBox);
assert.equal(reopenedPresentation.story(textBox.textStories[0].id).paragraphs
  .flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join(''), 'Native package API');
presentation.dispose();
reopenedPresentation.dispose();
const { initWasm, openWorkbook, exportXlsxMarkdown } = await import('@betteroffice/xlsx-native');
await initWasm();
const workbook = openWorkbook(xlsxBytes);
workbook.editCell(0, 1, 1, '=1+2');
workbook.setNumberFormat(0, 'B2', { type: 'custom', pattern: '0.000' });
workbook.patchRangeStyle(0, 'B2', {
  border: { preset: 'all', style: 'dashed', color: '#123456' },
  horizontalAlignment: 'right', textWrapping: 'wrap'
});
const proposal = workbook.propose('smoke-agent', null, [
  { sheet: 0, row: 1, col: 2, input: '12.5', numberFormat: 'currency' }
]);
assert.equal(proposal.cells[0].newText, '$12.50');
workbook.acceptProposal(proposal.id);
const target = { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2' } };
const reopenedWorkbook = openWorkbook(workbook.save());
assert.deepEqual(reopenedWorkbook.readCells({ ranges: [target] }).ranges[0].cells[0][0].value, { kind: 'number', value: 3 });
assert.ok((await exportXlsxMarkdown(xlsxBytes)).markdown.length > 0);
workbook.dispose();
reopenedWorkbook.dispose();
console.log('Native DOCX/XLSX/PPTX export names, package API, edit and save/reopen checks passed.');
