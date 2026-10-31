import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import * as wasm from '../../packages/pptx/src/index';
import { FORMATS } from '../../scripts/api-surface';
import { expectSavedPartsEqual } from '../native-parity';

type Pptx = typeof wasm;
type PresentationHandle = wasm.PresentationHandle;

const fixture = readFileSync(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));
let native: Pptx;

beforeAll(async () => {
  native = await import('./dist/index.js') as unknown as Pptx;
  await native.initWasm();
  await wasm.initWasm(readFileSync(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
});

async function compare(script: (presentation: PresentationHandle) => unknown) {
  const nativePresentation = native.openPresentation(fixture, { clientId: 4321 });
  const wasmPresentation = wasm.openPresentation(fixture, { clientId: 4321 });
  try {
    expect(script(nativePresentation)).toEqual(script(wasmPresentation));
    await expectSavedPartsEqual(nativePresentation.save(), wasmPresentation.save());
  } finally {
    nativePresentation.dispose();
    wasmPresentation.dispose();
  }
}

describe('pptx native/wasm package parity', () => {
  test('each entry exports the names of the matching @betteroffice/pptx entry', async () => {
    const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
    expect(Object.keys(FORMATS.pptx.entries)).toEqual(Object.keys(manifest.exports));
    for (const [entry, source] of Object.entries(FORMATS.pptx.entries)) {
      const nativeKeys = Object.keys(await import(new URL(manifest.exports[entry].import, import.meta.url).href));
      const wasmKeys = Object.keys(await import(`../../packages/pptx/${source}`));
      expect(nativeKeys.sort()).toEqual(wasmKeys.sort());
    }
  });

  test('text formatting and shape edits', async () => {
    await compare((presentation) => {
      const slideId = presentation.snapshot().slides[0].id;
      const textBox = presentation.addTextBox(slideId, {
        name: 'Parity text', rect: { x: 100000, y: 200000, width: 3000000, height: 1000000 },
        text: 'Initial', style: { fontFamily: 'Arial', fontSizePt: 18, color: '#123456' },
      });
      const shape = presentation.snapshot().slides[0].shapes.find((candidate) => candidate.id === textBox.shapeId)!;
      const storyId = shape.textStories[0].id;
      presentation.insertText(storyId, 7, ' styled', { bold: true, italic: true, color: '#654321', fontSizePt: 20 });
      presentation.formatText(storyId, 0, 7, { underline: 'sng', fontFamily: 'Arial', fontSizePt: 16 });
      presentation.setParagraphAlignment(storyId, 0, 14, 'ctr');
      presentation.insertParagraphBreak(storyId, 7);
      presentation.deleteText(storyId, 0, 2);
      const added = presentation.addShape(slideId, {
        name: 'Parity shape', geometry: 'roundRect',
        rect: { x: 500000, y: 600000, width: 2000000, height: 1500000 }, fill: '#112233',
      });
      presentation.setShapeFill(slideId, added.shapeId, '#abcdef');
      presentation.setShapeStroke(slideId, added.shapeId, { color: '#445566', widthPt: 2.5 });
      presentation.moveShape(slideId, added.shapeId, 700000, 800000);
      presentation.resizeShape(slideId, added.shapeId, 2500000, 1800000);
      presentation.setShapeRect(slideId, added.shapeId, { x: 900000, y: 1000000, width: 2600000, height: 1900000 });
      presentation.setShapeAdjust(slideId, added.shapeId, { adj: 25000 });
      const story = presentation.story(storyId);
      return {
        length: story.length,
        paragraphs: story.paragraphs.map(({ alignment, level, bulletJson, runs }) => ({ alignment, level, bulletJson, runs })),
      };
    });
  });

  test('content reads and structured exports', async () => {
    await compare((presentation) => {
      const read = presentation.readContent();
      if (!read.ok) throw new Error('Expected presentation content');
      const { version, ...content } = read;
      expect(typeof version).toBe('string');
      return content;
    });
    expect(await native.exportPptxMarkdown(fixture)).toEqual(await wasm.exportPptxMarkdown(fixture));
    expect(await native.exportPptxStructured(fixture)).toEqual(await wasm.exportPptxStructured(fixture));
  });
});
