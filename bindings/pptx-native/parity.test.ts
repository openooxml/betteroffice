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

describe('font pack diagnostics', () => {
  test('both packages report the installed packs alike', async () => {
    expect(await native.fontPackDiagnostics()).toEqual(await wasm.fontPackDiagnostics());
  });

  test('a broken installed pack is named with the same reason and detail by the native loader and the package check', async () => {
    const fs = require('node:fs') as typeof import('node:fs');
    const path = require('node:path') as typeof import('node:path');
    const { checkInstalledPacks } = await import('../../packages/fonts/src/packDiagnostics');
    const source = new URL('../../packages/fonts-intl', import.meta.url).pathname;
    const nativeReport = (directory: string) => {
      const script = `const raw = require('./index.js'); raw.configureFontPacks(undefined, ${JSON.stringify(directory)}, undefined); console.log('RESULT ' + raw.fontPackDiagnosticsJson());`;
      const child = Bun.spawnSync([process.execPath, '-e', script], { cwd: new URL('.', import.meta.url).pathname });
      const line = child.stdout.toString().split('\n').find((text) => text.startsWith('RESULT '));
      expect(line, child.stderr.toString()).toBeDefined();
      return JSON.parse(line!.slice(7));
    };
    const breakages: Array<[string, (assets: string) => void, string]> = [
      ['intact', () => {}, ''],
      ['missing file', (assets) => fs.rmSync(path.join(assets, 'NotoSansThai-Bold.ttf.zst')), 'missing-file'],
      ['other release', (assets) => {
        const manifest = path.join(assets, 'manifest.json');
        fs.writeFileSync(manifest, JSON.stringify({ ...JSON.parse(fs.readFileSync(manifest, 'utf8')), version: '0.0.1' }));
      }, 'version-mismatch'],
      ['no manifest', (assets) => fs.rmSync(path.join(assets, 'manifest.json')), 'missing-file'],
      ['unreadable manifest', (assets) => fs.writeFileSync(path.join(assets, 'manifest.json'), '{'), 'invalid-manifest'],
    ];
    for (const [name, damage, reason] of breakages) {
      const root = fs.mkdtempSync(path.join(require('node:os').tmpdir(), 'font-pack-'));
      try {
        fs.cpSync(path.join(source, 'src'), path.join(root, 'src'), { recursive: true });
        fs.cpSync(path.join(source, 'assets'), path.join(root, 'assets'), { recursive: true });
        damage(path.join(root, 'assets'));
        const { INTL_FONT_ASSET_URLS } = await import(path.join(root, 'src/index.ts'));
        const fromPackage = await checkInstalledPacks({
          latin: async () => undefined, intl: async () => INTL_FONT_ASSET_URLS, cjk: async () => undefined,
        });
        const fromNative = nativeReport(path.join(root, 'assets'));
        expect([name, fromNative.map((entry: { pack: string; reason: string }) => [entry.pack, entry.reason])])
          .toEqual([name, reason ? [['intl', reason]] : []]);
        expect([name, fromPackage]).toEqual([name, fromNative]);
      } finally {
        fs.rmSync(root, { recursive: true, force: true });
      }
    }
  });
});
