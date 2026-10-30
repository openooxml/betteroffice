import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const fixtureRoot = process.argv[2]
  ? pathToFileURL(`${resolve(process.argv[2])}/`)
  : new URL('../', import.meta.url);
const fixture = (path) => readFileSync(new URL(path, fixtureRoot));
const pngSignature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const docxBytes = fixture('apps/demo/public/betteroffice-demo.docx');
const xlsxBytes = fixture('crates/ooxml-opc/tests/fixtures/sample.xlsx');
const pptxBytes = fixture('apps/demo/public/betteroffice-demo.pptx');

for (const load of [(name) => import(name), async (name) => require(name)]) {
  const { openDocument } = await load('@betteroffice/docx-native');
  const document = await openDocument(docxBytes);
  const structure = await document.structure;
  assert.ok(structure.bodyParagraphs > 0);
  await document.registerFont({
    family: 'Calibri',
    data: fixture('crates/docx-raster/tests/assets/Carlito-Regular.ttf')
  });
  const layoutInput = JSON.parse(
    fixture('crates/docx-layout/tests/fixtures/single-page-multi-paragraph.input.json')
  );
  const page = await document.renderPage((await document.layout(layoutInput)).displayList);
  assert.deepEqual(page.data.subarray(0, 8), pngSignature);
  const reopenedDocument = await openDocument(await document.save());
  assert.equal((await reopenedDocument.structure).bodyParagraphs, structure.bodyParagraphs);

  const { openPresentation } = await load('@betteroffice/pptx-native');
  const presentation = await openPresentation(pptxBytes);
  const slideCount = await presentation.slideCount;
  assert.ok(slideCount > 0);
  await presentation.registerFont({
    family: 'Arial',
    data: fixture('crates/pptx-raster/tests/assets/Carlito-Regular.ttf')
  });
  const slide = await presentation.renderSlide(0);
  assert.deepEqual(slide.data.subarray(0, 8), pngSignature);
  const reopenedPresentation = await openPresentation(await presentation.save());
  assert.equal(await reopenedPresentation.slideCount, slideCount);
}
assert.ok(process.env.BETTEROFFICE_EXPORT_KEYS, 'BETTEROFFICE_EXPORT_KEYS names the export names of each wasm entry');
const exportKeys = JSON.parse(readFileSync(process.env.BETTEROFFICE_EXPORT_KEYS, 'utf8'));
for (const [format, entries] of Object.entries(exportKeys)) {
  for (const [entry, keys] of Object.entries(entries)) {
    const name = `@betteroffice/${format}-native${entry.slice(1)}`;
    assert.deepEqual(Object.keys(await import(name)).sort(), keys, name);
  }
}
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
console.log('Native DOCX/PPTX import, require, save/reopen and PNG checks and XLSX package API, recalculation and save/reopen checks passed.');
