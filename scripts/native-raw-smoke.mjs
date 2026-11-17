import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';

const format = process.argv[2];
const fixtures = {
  docx: 'apps/demo/public/betteroffice-demo.docx',
  pptx: 'apps/demo/public/betteroffice-demo.pptx',
  xlsx: 'crates/ooxml-opc/tests/fixtures/sample.xlsx'
};
assert.ok(Object.hasOwn(fixtures, format), `usage: node scripts/native-raw-smoke.mjs <${Object.keys(fixtures).join('|')}>`);
const bytes = new Uint8Array(readFileSync(new URL(`../${fixtures[format]}`, import.meta.url)));
const binding = createRequire(import.meta.url)(`../bindings/${format}-native/index.js`);

if (format === 'xlsx') {
  const document = binding.XlsxDocument.open(bytes);
  assert.ok(document.sheetCount() > 0);
  const reopened = binding.XlsxDocument.open(document.saveBytes());
  assert.equal(reopened.sheetCount(), document.sheetCount());
} else if (format === 'pptx') {
  const document = binding.PptxDocument.openReplayBaseline(bytes);
  const slides = JSON.parse(document.snapshotJson()).slides;
  assert.ok(slides.length > 0);
  const reopened = binding.PptxDocument.openReplayBaseline(document.saveBytes());
  assert.equal(JSON.parse(reopened.snapshotJson()).slides.length, slides.length);
} else {
  assert.ok(Array.isArray(JSON.parse(binding.font_preflight(bytes)).faces));
}
console.log(`${format}-native loaded and ran on ${process.platform} ${process.arch}`);
