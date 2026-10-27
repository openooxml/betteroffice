import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import { transformSync } from 'esbuild';
import { Window } from 'happy-dom';
import JSZip from 'jszip';
import * as manifest from '../../packages/fonts/src/manifest';
import { fontProvider } from '../../packages/fonts/src/provider';
import { normalFontIndex } from './xlsx-styles';

interface RendererFace {
  family: string;
  bold: boolean;
  italic: boolean;
  bytes: Uint8Array;
}

interface BrowserFace {
  family: string;
  file: string;
  weight: string;
  style: string;
}

async function loadHarness() {
  const source = await readFile(new URL('./harness.ts', import.meta.url), 'utf8');
  const code = transformSync(
    `${source}\nglobalThis.fontFunctions = { fontsFor, fallbackFontsFor };`,
    { loader: 'ts', format: 'cjs' }
  ).code;
  const registrations: BrowserFace[] = [];
  const provider = fontProvider(async (face) => new TextEncoder().encode(face.file).buffer);
  const modules: Record<string, unknown> = {
    jszip: JSZip,
    '../../packages/fonts/src/cdn': { createFontProvider: () => provider },
    '../../packages/fonts/src/manifest': manifest,
    './xlsx-styles': { normalFontIndex },
  };
  const context = {
    window: {},
    location: { search: '?format=pptx' },
    URLSearchParams,
    DOMParser: new Window().DOMParser,
    document: {
      fonts: {
        add: (face: BrowserFace) => registrations.push(face),
        delete: (face: BrowserFace) => registrations.splice(registrations.indexOf(face), 1),
      },
    },
    FontFace: class {
      family: string;
      file: string;
      weight: string;
      style: string;

      constructor(family: string, bytes: ArrayBuffer, descriptors: FontFaceDescriptors) {
        this.family = family;
        this.file = new TextDecoder().decode(bytes);
        this.weight = descriptors.weight ?? '400';
        this.style = descriptors.style ?? 'normal';
      }

      async load() {
        return this;
      }
    },
    require: (id: string) => {
      if (!(id in modules)) throw new Error(`Unexpected harness import: ${id}`);
      return modules[id];
    },
    fontFunctions: undefined as unknown as {
      fontsFor(bytes: Uint8Array): Promise<RendererFace[]>;
      fallbackFontsFor(bytes: Uint8Array): Promise<RendererFace[]>;
    },
  };
  runInNewContext(code, context);
  return { ...context.fontFunctions, registrations };
}

async function deck(text: string, family?: string) {
  const zip = new JSZip();
  zip.file(
    'ppt/slides/slide1.xml',
    `<root>${family ? `<latin typeface="${family}"/>` : ''}<r lang="zh-CN"><t>${text}</t></r></root>`
  );
  return zip.generateAsync({ type: 'uint8array' });
}

function rendererFaces(faces: RendererFace[]) {
  return faces.map(({ bytes, ...face }) => ({ ...face, file: new TextDecoder().decode(bytes) }));
}

test('registers a Regular-only CJK script fallback once at 400 and preserves renderer bold flags', async () => {
  const harness = await loadHarness();
  const faces = await harness.fallbackFontsFor(await deck('中文'));
  expect(harness.registrations).toEqual([
    { family: 'Noto Sans SC', file: 'NotoSansSC-Regular.otf', weight: '400', style: 'normal' },
  ]);
  expect(rendererFaces(faces)).toEqual([
    { family: 'Noto Sans SC', bold: false, italic: false, file: 'NotoSansSC-Regular.otf' },
    { family: 'Noto Sans SC', bold: true, italic: false, file: 'NotoSansSC-Regular.otf' },
  ]);
});

test('registers a named Regular-only CJK family once without bold or italic aliases', async () => {
  const harness = await loadHarness();
  const faces = await harness.fontsFor(await deck('中文', 'Microsoft YaHei'));
  expect(harness.registrations.filter((face) => face.family === 'Microsoft YaHei')).toEqual([
    { family: 'Microsoft YaHei', file: 'NotoSansSC-Regular.otf', weight: '400', style: 'normal' },
  ]);
  expect(rendererFaces(faces.filter((face) => face.family === 'Microsoft YaHei'))).toEqual([
    { family: 'Microsoft YaHei', bold: false, italic: false, file: 'NotoSansSC-Regular.otf' },
    { family: 'Microsoft YaHei', bold: true, italic: false, file: 'NotoSansSC-Regular.otf' },
    { family: 'Microsoft YaHei', bold: false, italic: true, file: 'NotoSansSC-Regular.otf' },
    { family: 'Microsoft YaHei', bold: true, italic: true, file: 'NotoSansSC-Regular.otf' },
  ]);
});

test('registers real bundled bold and italic files with their metadata', async () => {
  const harness = await loadHarness();
  const faces = await harness.fontsFor(await deck('Hello'));
  expect(harness.registrations.filter((face) => face.family === 'Arial')).toEqual([
    { family: 'Arial', file: 'LiberationSans-Regular.ttf', weight: '400', style: 'normal' },
    { family: 'Arial', file: 'LiberationSans-Bold.ttf', weight: '700', style: 'normal' },
    { family: 'Arial', file: 'LiberationSans-Italic.ttf', weight: '400', style: 'italic' },
    { family: 'Arial', file: 'LiberationSans-BoldItalic.ttf', weight: '700', style: 'italic' },
  ]);
  expect(rendererFaces(faces.filter((face) => face.family === 'Arial'))).toEqual([
    { family: 'Arial', bold: false, italic: false, file: 'LiberationSans-Regular.ttf' },
    { family: 'Arial', bold: true, italic: false, file: 'LiberationSans-Bold.ttf' },
    { family: 'Arial', bold: false, italic: true, file: 'LiberationSans-Italic.ttf' },
    { family: 'Arial', bold: true, italic: true, file: 'LiberationSans-BoldItalic.ttf' },
  ]);
});

test('registers a script fallback with a real bundled bold file at 700', async () => {
  const harness = await loadHarness();
  const faces = await harness.fallbackFontsFor(await deck('مرحبا'));
  expect(harness.registrations).toEqual([
    { family: 'Noto Sans Arabic', file: 'NotoSansArabic-Regular.ttf', weight: '400', style: 'normal' },
    { family: 'Noto Sans Arabic', file: 'NotoSansArabic-Bold.ttf', weight: '700', style: 'normal' },
  ]);
  expect(rendererFaces(faces)).toEqual([
    { family: 'Noto Sans Arabic', bold: false, italic: false, file: 'NotoSansArabic-Regular.ttf' },
    { family: 'Noto Sans Arabic', bold: true, italic: false, file: 'NotoSansArabic-Bold.ttf' },
  ]);
});

test('italic Segoe UI registers its true italic last resort beside upright Selawik', async () => {
  const harness = await loadHarness();
  const zip = new JSZip();
  zip.file('ppt/slides/slide1.xml', '<root><latin typeface="Segoe UI"/><r><t>Text</t></r></root>');
  const faces = rendererFaces(await harness.fontsFor(await zip.generateAsync({ type: 'uint8array' })));
  expect(faces.filter((face) => face.family === 'Segoe UI')).toEqual([
    { family: 'Segoe UI', bold: false, italic: false, file: 'Selawik-Regular.ttf' },
    { family: 'Segoe UI', bold: true, italic: false, file: 'Selawik-Bold.ttf' },
    { family: 'Segoe UI', bold: false, italic: true, file: 'Carlito-Italic.ttf' },
    { family: 'Segoe UI', bold: true, italic: true, file: 'Carlito-BoldItalic.ttf' },
  ]);
  expect(harness.registrations.filter((face) => face.family === 'Segoe UI')).toEqual([
    { family: 'Segoe UI', file: 'Selawik-Regular.ttf', weight: '400', style: 'normal' },
    { family: 'Segoe UI', file: 'Selawik-Bold.ttf', weight: '700', style: 'normal' },
    { family: 'Segoe UI', file: 'Carlito-Italic.ttf', weight: '400', style: 'italic' },
    { family: 'Segoe UI', file: 'Carlito-BoldItalic.ttf', weight: '700', style: 'italic' },
  ]);
});

for (const order of ['named first', 'fallback first'] as const) {
  test(`a real script face replaces every last-resort alias under its family (${order})`, async () => {
    const harness = await loadHarness();
    const bytes = await deck('中文', 'Noto Sans SC');
    if (order === 'named first') {
      const named = await harness.fontsFor(bytes);
      expect(rendererFaces(named.filter((face) => face.family === 'Noto Sans SC'))[0].file).not.toBe(
        'NotoSansSC-Regular.otf'
      );
      await harness.fallbackFontsFor(bytes);
    } else {
      await harness.fallbackFontsFor(bytes);
      await harness.fontsFor(bytes);
    }
    expect(harness.registrations.filter((face) => face.family === 'Noto Sans SC')).toEqual([
      { family: 'Noto Sans SC', file: 'NotoSansSC-Regular.otf', weight: '400', style: 'normal' },
    ]);
  });
}
