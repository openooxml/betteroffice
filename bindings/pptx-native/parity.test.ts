// Normalization: ZIP container headers, entry order, directory entries and compression only.
// These can differ between Rust targets; every part name and uncompressed byte must match.
// Both sides use clientId=4321 and the wasm boundary's author='wasm', origin='local'.
import { readFileSync } from 'node:fs';
import { beforeAll, expect, test } from 'bun:test';
import { initWasm, openPresentation as openWasm } from '../../packages/pptx/src/wasm/loader';
import { expectSavedPartsEqual } from '../native-parity';
import { openPresentation } from './index.js';

const fixture = readFileSync(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));

beforeAll(async () => {
  await initWasm(readFileSync(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
});

test('pptx native/wasm text and shape saved parity', async () => {
  const native = await openPresentation(fixture, { clientId: 4321, author: 'wasm', origin: 'local' });
  const wasm = openWasm(fixture, { clientId: 4321 });
  async function script(presentation: typeof native | typeof wasm) {
    const slideId = (await presentation.snapshot()).slides[0].id;
    const textBox = await presentation.addTextBox(slideId, {
      name: 'Parity text', rect: { x: 100000, y: 200000, width: 3000000, height: 1000000 },
      text: 'Initial', style: { fontFamily: 'Arial', fontSizePt: 18, color: '#123456' }
    });
    const snapshot = await presentation.snapshot();
    const shape = snapshot.slides[0].shapes.find((candidate) => candidate.id === textBox.shapeId)!;
    const storyId = shape.textStories[0].id;
    await presentation.insertText(storyId, 7, ' styled', { bold: true, italic: true, color: '#654321', fontSizePt: 20 });
    await presentation.formatText(storyId, 0, 7, { underline: 'sng', fontFamily: 'Arial', fontSizePt: 16 });
    await presentation.setParagraphAlignment(storyId, 0, 14, 'ctr');
    await presentation.insertParagraphBreak(storyId, 7);
    await presentation.deleteText(storyId, 0, 2);
    const added = await presentation.addShape(slideId, {
      name: 'Parity shape', geometry: 'roundRect',
      rect: { x: 500000, y: 600000, width: 2000000, height: 1500000 }, fill: '#112233'
    });
    await presentation.setShapeFill(slideId, added.shapeId, '#abcdef');
    await presentation.setShapeStroke(slideId, added.shapeId, { color: '#445566', widthPt: 2.5 });
    await presentation.moveShape(slideId, added.shapeId, 700000, 800000);
    await presentation.resizeShape(slideId, added.shapeId, 2500000, 1800000);
    await presentation.setShapeRect(slideId, added.shapeId, { x: 900000, y: 1000000, width: 2600000, height: 1900000 });
    await presentation.setShapeAdjust(slideId, added.shapeId, { adj: 25000 });
    return presentation.story(storyId);
  }
  try {
    expect(await script(native)).toEqual(await script(wasm));
    await expectSavedPartsEqual(await native.save(), wasm.save());
  } finally {
    wasm.dispose();
  }
});
