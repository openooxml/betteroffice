import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, spyOn, test } from 'bun:test';
import { copyDecodedFace, fontBytes } from '../../shared/office-session/fontDecoder';

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
    expect(manifest.dependencies).toEqual({
      ...wasmManifest.dependencies,
      '@betteroffice/fonts-cjk': 'workspace:^',
      '@betteroffice/fonts-intl': 'workspace:^',
      '@betteroffice/fonts-latin': 'workspace:^',
    });
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

  test('loads every font pack from the installed packages without fetching', async () => {
    const fetch = spyOn(globalThis, 'fetch').mockRejectedValue(new Error('Network disabled'));
    try {
      const fonts = await import(Bun.resolveSync('@betteroffice/fonts', fileURLToPath(new URL('.', import.meta.url)))) as
        typeof import('../../packages/fonts/src/index');
      const { decodeFontFace } = await import('./js/engine');
      const source = fonts.createFontProvider(undefined, {
        decoder: async (bytes, { codec, dict, maxOutput }) =>
          copyDecodedFace(decodeFontFace(fontBytes(bytes), codec, dict === undefined ? undefined : fontBytes(dict), maxOutput)),
      });
      for (const family of ['Carlito', 'Noto Sans Devanagari', 'Noto Sans JP']) {
        const face = fonts.BUNDLED_FONTS.find((face) => face.family === family && face.weight === 400 && face.style === 'normal')!;
        const bytes = await (source.resolve(family, false, false) ?? source.resolveFamily(family, false, false))!();
        expect(bytes.byteLength).toBe(face.byteLength);
      }
      expect(fetch).not.toHaveBeenCalled();
    } finally {
      fetch.mockRestore();
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
