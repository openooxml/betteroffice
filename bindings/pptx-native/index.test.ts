import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, spyOn, test } from 'bun:test';

const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
const wasmManifest = JSON.parse(readFileSync(new URL('../../packages/pptx/package.json', import.meta.url), 'utf8'));
const fixture = readFileSync(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));
const dist = new URL('./dist/', import.meta.url);

async function entry(name: string) {
  return import(new URL(manifest.exports[name].import, import.meta.url).href);
}

describe('@betteroffice/pptx-native', () => {
  test('exposes the entries of @betteroffice/pptx and refuses browsers', () => {
    expect(Object.keys(manifest.exports)).toEqual(Object.keys(wasmManifest.exports));
    for (const [name, conditions] of Object.entries(wasmManifest.exports) as [string, Record<string, string>][]) {
      expect(manifest.exports[name]).toEqual({
        types: conditions.types,
        browser: './dist/browser.js',
        import: conditions.import,
      });
    }
    expect(manifest.dependencies).toEqual(wasmManifest.dependencies);
    expect(manifest.sideEffects).toEqual(wasmManifest.sideEffects);
    expect(JSON.parse(readFileSync(new URL('package.json', dist), 'utf8'))).toEqual({ type: 'module' });
  });

  test('the browser entry throws', async () => {
    await expect(import(new URL('browser.js', dist).href)).rejects.toThrow(
      'use @betteroffice/pptx in browsers'
    );
  });

  test('ships no wasm engine', () => {
    for (const file of readdirSync(dist)) {
      expect(file.endsWith('.wasm')).toBe(false);
      const text = readFileSync(new URL(file, dist), 'utf8');
      expect(text).not.toContain('__wbindgen');
      expect(text).not.toContain('_bg.wasm');
    }
  });

  test('opens, edits, saves and reopens on the native engine', async () => {
    const instantiate = spyOn(WebAssembly, 'instantiate');
    const compile = spyOn(WebAssembly, 'compile');
    try {
      const pptx = await entry('.') as typeof import('../../packages/pptx/src/index');
      await pptx.initWasm();
      const presentation = pptx.openPresentation(fixture);
      try {
        const snapshot = presentation.snapshot();
        expect(snapshot.slides.length).toBeGreaterThan(0);
        const added = presentation.addTextBox(snapshot.slides[0].id, {
          name: 'Native API', rect: { x: 100000, y: 200000, width: 3000000, height: 1000000 },
          text: 'Native', style: { fontFamily: 'Arial', fontSizePt: 18 },
        });
        const shape = presentation.snapshot().slides[0].shapes.find(shape => shape.id === added.shapeId)!;
        presentation.insertText(shape.textStories[0].id, 6, ' package API');
        const reopened = pptx.openPresentation(presentation.save());
        try {
          const saved = reopened.snapshot();
          expect(saved.slides.length).toBe(snapshot.slides.length);
          const textBox = saved.slides[0].shapes.find(shape => shape.name === 'Native API')!;
          const story = reopened.story(textBox.textStories[0].id);
          expect(story.paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.text)).join(''))
            .toBe('Native package API');
        } finally {
          reopened.dispose();
        }
      } finally {
        presentation.dispose();
      }
      expect(instantiate).not.toHaveBeenCalled();
      expect(compile).not.toHaveBeenCalled();
    } finally {
      instantiate.mockRestore();
      compile.mockRestore();
    }
  });

  test('runs a worker session on the native engine', async () => {
    const { openPresentationSession } = await entry('./worker') as typeof import('../../packages/pptx/src/worker');
    const session = await openPresentationSession(fixture);
    try {
      expect(session.state.slides.length).toBeGreaterThan(0);
      expect(await session.call.slides()).toEqual(session.state.slides);
      expect((await session.call.readContent()).ok).toBe(true);
      expect((await session.save()).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    } finally {
      await session.dispose();
    }
  });
});
