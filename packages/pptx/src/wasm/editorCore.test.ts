import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { exportPptxStructured, initWasm, openPresentation, renderPptxMarkdown } from '../index';
import { sessionWasmAssetUrl, wasmAssetUrl } from './asset';
import { engineKind } from './engine';
import { compileSlideJson } from './generated/pptx_wasm.js';

const root = resolve(import.meta.dir, '../../../..');
let deck: Uint8Array;
let font: Uint8Array;

test('the session core ships next to the full core', () => {
  expect(new URL('pptx_wasm_lean_bg.wasm', wasmAssetUrl()).href).toBe(sessionWasmAssetUrl().href);
});

describe.skipIf(engineKind === 'native')('editor core', () => {
  beforeAll(async () => {
    const [wasm, pptx, face] = await Promise.all([
      readFile(resolve(import.meta.dir, 'generated/pptx_wasm_lean_bg.wasm')),
      readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
      readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
    ]);
    await initWasm(wasm);
    deck = pptx;
    font = face;
  });

  test('opens, lays out, edits and saves', () => {
    const handle = openPresentation(deck, { clientId: 9301, fonts: [{ family: 'Liberation Sans', bytes: font }] });
    try {
      expect(handle.layoutSlide(0).primitives.length).toBeGreaterThan(0);
      const story = handle.snapshot().slides[0].shapes.find((shape) => shape.textStories.length > 0)!.textStories[0];
      handle.insertText(story.id, 0, 'core ');
      const reopened = openPresentation(handle.save(), { clientId: 9302 });
      try {
        const text = reopened.snapshot().slides[0].shapes.flatMap((shape) => shape.textStories)
          .flatMap((entry) => entry.paragraphs).flatMap((paragraph) => paragraph.runs).map((run) => run.text).join('');
        expect(text).toContain('core ');
      } finally {
        reopened.dispose();
      }
    } finally {
      handle.dispose();
    }
  });

  test('exports structured content from bytes', async () => {
    const structured = await exportPptxStructured(deck);
    expect(structured.slides.length).toBeGreaterThan(0);
  });

  test('refuses the one-shot slide compiler and Markdown renderer', async () => {
    const structured = await exportPptxStructured(deck);
    await expect(renderPptxMarkdown(structured)).rejects.toThrow('not available in the editor core');
    expect(() => compileSlideJson('{}')).toThrow('not available in the editor core');
  });
});
