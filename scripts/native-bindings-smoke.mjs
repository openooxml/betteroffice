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

  const { openWorkbook } = await load('@betteroffice/xlsx-native');
  const workbook = await openWorkbook(xlsxBytes);
  const edit = workbook.set(0, 'B2', '=1+2');
  const value = workbook.value(0, 'B2');
  await edit;
  assert.equal((await value).number, 3);
  await workbook.setNumberFormat(0, 'B2', 'plainText');
  await workbook.setNumberFormat(0, 'B2', { type: 'custom', pattern: '0.000' });
  await workbook.setNumberFormat(0, 'B2', 'increaseDecimal');
  await workbook.setNumberFormat(0, 'B2', 'decreaseDecimal');
  await workbook.setStyle(0, 'B2', {
    border: { preset: 'all', style: 'dashed', color: '#123456' },
    horizontalAlignment: 'right', textWrapping: 'wrap'
  });
  await workbook.setStyle(0, 'B2', { clear: ['borders', 'textWrapping'] });
  const proposal = await workbook.propose({
    agentId: 'smoke-agent',
    edits: [{ sheet: 0, address: 'C2', input: '12.5', numberFormat: 'currency' }]
  });
  assert.equal(proposal.edits[0].after, '$12.50');
  await workbook.acceptProposal(proposal.id);
  const sheet = await workbook.renderSheet({ sheet: 0, range: 'A1:B3' });
  assert.deepEqual(sheet.data.subarray(0, 8), pngSignature);
  const reopenedWorkbook = await openWorkbook(await workbook.save());
  assert.equal((await reopenedWorkbook.value(0, 'B2')).number, 3);

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
console.log('Native DOCX/PPTX/XLSX import, require, save/reopen, PNG and recalculation checks passed.');
