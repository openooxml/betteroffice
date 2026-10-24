import { expect, test } from 'bun:test';
import type { DisplayPage, GlyphRunPrimitive } from '../layout/render/displayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { projectCommandEchoContext, type EchoGlyphProvider } from './residentCommandViewProjection';
import { indexEchoGlyphSet, placeEchoGlyphs } from './residentEchoGlyphs';

function input(text = 'é🙂a') {
  const run: GlyphRunPrimitive = { kind: 'glyphRun', text, size: 12, fontId: 1, color: '#000',
    fallbackFont: '12px Arial', docStart: 1, docEnd: text.length + 1,
    glyphs: [{ id: 1, cluster: 0, x: 0, y: 14, advance: 10 },
      { id: 2, cluster: 2, x: 10, y: 14, advance: 10 },
      { id: 3, cluster: 6, x: 20, y: 14, advance: 10 }] };
  const page: DisplayPage = { pageIndex: 0, width: 100, height: 40, primitives: [
    { kind: 'rect', x: 0, y: 0, w: 100, h: 20, fill: '#ccc' }, run,
    { kind: 'image', relId: 'media', x: 50, y: 0, w: 20, h: 20 },
  ] };
  const frame: RetainedFrame = { protocolVersion: 1, docEpoch: 1, layoutEpoch: 1, frameEpoch: 3,
    pages: [{ pageId: 1n, pageIndex: 0, fingerprint: 1n,
      primitiveIds: new BigUint64Array([10n, 11n, 12n]), page }],
    damagedPageIds: new Set([1n]), removedPageIds: new Set(), displayList: { pages: [page] } };
  const engine = {
    proposalEngine: { readParagraphs: () => ({ ok: true, paragraphs: [{ paraId: 'p', text, atoms: [] }] }) },
    geometryReader: { locateParagraph: () => ({ start: 0, end: text.length }) },
    residentCaretSnapshot: () => ({ frameEpoch: 3,
      caretRect: { pageId: '1', pageIndex: 0, x: 20, y: 2, height: 12 } }),
  } as unknown as DocxWorkerCommandEngine;
  const at = { story: 'body', paraId: 'p', offset: 3 };
  return { engine, selection: { anchor: at, focus: at }, layout: { frame, layoutRevision: 2 }, run };
}

test('d2c_projection_maps_byte_clusters_to_utf16_graphemes', () => {
  const source = input();
  const context = projectCommandEchoContext(source);
  expect(context?.boundaries.map(({ utf16, grapheme, caret }) => [utf16, grapheme, caret.x]))
    .toEqual([[0, 0, 0], [1, 1, 10], [3, 2, 20], [4, 3, 30]]);
  expect(context?.selection).toEqual({ anchorGrapheme: 2, focusGrapheme: 2 });
  expect(context?.damage[0]?.omittedPrimitiveIds).toEqual(['11']);
  expect(context?.serializedBytes).toBe(new TextEncoder().encode(JSON.stringify(context)).byteLength);
  expect(projectCommandEchoContext({ ...source, selection: null })).toBeUndefined();
});

test('retained caret projection declines mismatched epochs and missing or unbuilt pages', () => {
  const source = input();
  const at = { ...source.selection.focus, offset: 1 };
  source.selection = { anchor: at, focus: at };
  source.engine.residentCaretSnapshot = () => ({ frameEpoch: 3,
    caretRect: { pageId: '1', pageIndex: 0, x: 10, y: 2, height: 12 } });
  expect(projectCommandEchoContext(source)?.selection).toEqual({ anchorGrapheme: 1, focusGrapheme: 1 });
  const frame = source.layout.frame;
  source.layout.frame = { ...frame, frameEpoch: 4 };
  expect(projectCommandEchoContext(source)).toBeUndefined();
  source.layout.frame = { ...frame, pages: [{ ...frame.pages[0]!, pageId: 2n }] };
  expect(projectCommandEchoContext(source)).toBeUndefined();
  source.layout.frame = { ...frame, pages: [] };
  expect(projectCommandEchoContext(source)).toBeUndefined();
  source.layout.frame = { ...frame, pages: [{ ...frame.pages[0]!, page: { ...frame.pages[0]!.page, unbuilt: true } }] };
  expect(projectCommandEchoContext(source)).toBeUndefined();
});

test('d2c_projection_declines_oversize_or_unsupported_context', () => {
  const source = input();
  source.run.rtl = true;
  expect(projectCommandEchoContext(source)).toBeUndefined();
  expect(projectCommandEchoContext(input('x'.repeat(16 * 1024 + 1)))).toBeUndefined();
});

test('d2c_projection_reads_only_the_active_paragraph_and_caret_page', () => {
  const source = input();
  const reads: unknown[] = [];
  Object.assign(source.engine, {
    paragraphs() { throw new Error('whole body read'); },
    storySegments() { throw new Error('whole story read'); },
  });
  Object.assign(source.engine.geometryReader, {
    storySegments() { throw new Error('whole story projection'); },
    positionOutline() { throw new Error('whole story outline'); },
  });
  const readParagraphs = source.engine.proposalEngine.readParagraphs;
  source.engine.proposalEngine.readParagraphs = (request) => {
    reads.push(request);
    return readParagraphs(request);
  };
  const retained = source.layout.frame.pages[0]!;
  source.layout.frame = { ...source.layout.frame, pages: [retained, { ...retained, pageId: 2n, pageIndex: 1,
    page: { ...retained.page, get primitives(): never { throw new Error('offscreen page read'); } } }] };
  expect(projectCommandEchoContext(source)?.text).toBe('é🙂a');
  expect(reads).toEqual([{ story: 'body', paraIds: ['p'], view: 'accepted' }]);
});

test('projection echoes the caret line of a wrapped paragraph', () => {
  const line = (text: string, docStart: number, y: number): GlyphRunPrimitive => ({ kind: 'glyphRun', text, size: 12,
    fontId: 1, color: '#000', fallbackFont: '12px Arial', docStart, docEnd: docStart + text.length,
    glyphs: [...text].map((_, index) => ({ id: index + 1, cluster: index, x: index * 10, y, advance: 10 })) });
  const page: DisplayPage = { pageIndex: 0, width: 100, height: 60, primitives: [line('ab ', 1, 14), line('cd', 4, 34)] };
  const frame: RetainedFrame = { protocolVersion: 1, docEpoch: 1, layoutEpoch: 1, frameEpoch: 3,
    pages: [{ pageId: 1n, pageIndex: 0, fingerprint: 1n, primitiveIds: new BigUint64Array([10n, 11n]), page }],
    damagedPageIds: new Set([1n]), removedPageIds: new Set(), displayList: { pages: [page] } };
  const engine = {
    proposalEngine: { readParagraphs: () => ({ ok: true, paragraphs: [{ paraId: 'p', text: 'ab cd', atoms: [] }] }) },
    geometryReader: { locateParagraph: () => ({ start: 0, end: 5 }) },
    residentCaretSnapshot: () => ({ frameEpoch: 3, caretRect: { pageId: '1', pageIndex: 0, x: 10, y: 22, height: 12 } }),
  } as unknown as DocxWorkerCommandEngine;
  const at = { story: 'body', paraId: 'p', offset: 4 };
  const context = projectCommandEchoContext({ engine, selection: { anchor: at, focus: at }, layout: { frame, layoutRevision: 2 } });
  expect(context?.text).toBe('cd');
  expect(context?.selection).toEqual({ anchorGrapheme: 1, focusGrapheme: 1 });
  expect(context?.boundaries.map(({ utf16, caret }) => [utf16, caret.x])).toEqual([[0, 0], [1, 10], [2, 20]]);
  expect(context?.styleRuns.map(({ baseline }) => baseline)).toEqual([{ x: 0, y: 34 }]);
  expect(context?.damage[0]).toMatchObject({ rect: { x: 0, y: 22, width: 100, height: 18 }, omittedPrimitiveIds: ['11'] });
});

test('projection bounds its scan by the caret line, not by the glyphs on the rest of the page', () => {
  const source = input();
  const other = (index: number): GlyphRunPrimitive => ({ kind: 'glyphRun', text: 'y'.repeat(60), size: 12, fontId: 1,
    color: '#000', fallbackFont: '12px Arial', docStart: 100 + index * 61, docEnd: 160 + index * 61,
    glyphs: Array.from({ length: 60 }, (_, glyph) => ({ id: glyph + 1, cluster: glyph, x: glyph, y: 40 + index * 20, advance: 1 })) });
  const page = source.layout.frame.pages[0]!.page;
  const others = Array.from({ length: 20 }, (_, index) => other(index));
  const extended = { ...page, height: 500, primitives: [...page.primitives, ...others] };
  source.layout.frame = { ...source.layout.frame, pages: [{ ...source.layout.frame.pages[0]!, page: extended,
    primitiveIds: new BigUint64Array([10n, 11n, 12n, ...others.map((_, index) => BigInt(100 + index))]) }],
    displayList: { pages: [extended] } };
  expect(others.reduce((total, run) => total + run.glyphs.length, 0)).toBeGreaterThan(1024);
  expect(projectCommandEchoContext(source)?.text).toBe('é🙂a');
});

test('projection context does not grow with the primitives on the rest of the page', () => {
  const source = input();
  const small = projectCommandEchoContext(source)!;
  const page = source.layout.frame.pages[0]!.page;
  const shading = Array.from({ length: 2048 }, (_, index) => ({ kind: 'rect' as const, x: 0, y: 30, w: 1, h: 1, fill: `#${index}` }));
  const extended = { ...page, primitives: [...page.primitives, ...shading] };
  source.layout.frame = { ...source.layout.frame, pages: [{ ...source.layout.frame.pages[0]!, page: extended,
    primitiveIds: new BigUint64Array([10n, 11n, 12n, ...shading.map((_, index) => BigInt(1000 + index))]) }],
    displayList: { pages: [extended] } };
  const large = projectCommandEchoContext(source);
  expect(large?.damage[0]?.omittedPrimitiveIds).toEqual(['11']);
  expect(large?.serializedBytes).toBe(small.serializedBytes);
});

test('projection echoes a cluster-positioned line as one style run within the context cap', () => {
  const text = 'a'.repeat(160);
  const runs: GlyphRunPrimitive[] = [...text].map((character, index) => ({ kind: 'glyphRun', text: character, size: 12,
    fontId: 1, color: '#000', fallbackFont: '12px Arial', docStart: 1 + index, docEnd: 2 + index,
    glyphs: [{ id: 1, cluster: 0, x: index * 5, y: 14, advance: 5 }] }));
  const page: DisplayPage = { pageIndex: 0, width: 1000, height: 40, primitives: runs };
  const frame: RetainedFrame = { protocolVersion: 1, docEpoch: 1, layoutEpoch: 1, frameEpoch: 3,
    pages: [{ pageId: 1n, pageIndex: 0, fingerprint: 1n, primitiveIds: new BigUint64Array(runs.map((_, index) => BigInt(index + 1))), page }],
    damagedPageIds: new Set([1n]), removedPageIds: new Set(), displayList: { pages: [page] } };
  const engine = {
    proposalEngine: { readParagraphs: () => ({ ok: true, paragraphs: [{ paraId: 'p', text, atoms: [] }] }) },
    geometryReader: { locateParagraph: () => ({ start: 0, end: text.length }) },
    residentCaretSnapshot: () => ({ frameEpoch: 3, caretRect: { pageId: '1', pageIndex: 0, x: 400, y: 2, height: 12 } }),
  } as unknown as DocxWorkerCommandEngine;
  const at = { story: 'body', paraId: 'p', offset: 80 };
  const context = projectCommandEchoContext({ engine, selection: { anchor: at, focus: at }, layout: { frame, layoutRevision: 2 } });
  expect(context?.text).toBe(text);
  expect(context?.styleRuns.map(({ startUtf16, endUtf16, baseline }) => [startUtf16, endUtf16, baseline.x])).toEqual([[0, 160, 0]]);
  expect(context!.serializedBytes).toBeLessThanOrEqual(16 * 1024);
});

type GlyphLineOptions = { shiftX?: number; caretLine?: number; later?: DisplayPage['primitives'];
  x?: number; advance?: number; kerned?: boolean; kerning?: number; right?: number; lines?: string[] };

function glyphLine(options: GlyphLineOptions = {}) {
  const lines = options.lines ?? ['ab ', 'cd'];
  const text = lines.join('');
  const runs: GlyphRunPrimitive[] = [];
  let doc = 1;
  lines.forEach((line, lineIndex) => {
    let x = options.x ?? 96;
    [...line].forEach((char, index) => {
      const kern = options.kerned && char === 'A' && line[index + 1] === 'V' ? -2 : 0;
      runs.push({ kind: 'glyphRun', text: char, size: 13.333, fontId: 1, color: '#000000',
        echo: { x: 96, right: options.right ?? 720,
          kern: (options.kerning ?? 0) > 0 && (options.kerning ?? 0) <= 20,
          kerning: (options.kerning ?? 0) / 2, primaryFont: true },
        fallbackFont: '400 13.333px Calibri, sans-serif', docStart: doc, docEnd: doc + 1,
        glyphs: [{ id: char.charCodeAt(0), cluster: 0, y: 120 + lineIndex * 20, advance: 10 + kern,
          x: x + (options.shiftX && lineIndex === 1 && index === 1 ? options.shiftX : 0) }] });
      x += (options.advance ?? 10) + kern;
      doc += 1;
    });
  });
  const primitives = [...runs, ...(options.later ?? [])];
  const page: DisplayPage = { pageIndex: 0, width: 816, height: 1056, primitives,
    contentBounds: { x: 96, y: 96, width: options.right === undefined ? 624 : options.right - 96, height: 864 } };
  const frame: RetainedFrame = { protocolVersion: 1, docEpoch: 1, layoutEpoch: 1, frameEpoch: 3,
    pages: [{ pageId: 1n, pageIndex: 0, fingerprint: 1n,
      primitiveIds: new BigUint64Array(primitives.map((_, index) => BigInt(index + 1))), page }],
    damagedPageIds: new Set([1n]), removedPageIds: new Set(), displayList: { pages: [page] } };
  const caretLine = options.caretLine ?? lines.length - 1;
  const lineStart = lines.slice(0, caretLine).join('').length;
  const caretGlyph = runs[lineStart + Math.min(1, lines[caretLine]!.length - 1)]!.glyphs[0]!;
  const engine = {
    proposalEngine: { readParagraphs: () => ({ ok: true, paragraphs: [{ paraId: 'p', text, atoms: [] }] }) },
    geometryReader: { locateParagraph: () => ({ start: 0, end: text.length }), paragraphSpans: () => [{ paraId: 'p' }] },
    paragraphSegments: () => [[{ kind: 'text', text,
      attributes: options.kerning ? { characterSpacing: { kerning: options.kerning } } : {} }]],
    residentCaretSnapshot: () => ({ frameEpoch: 3,
      caretRect: { pageId: '1', pageIndex: 0, x: caretGlyph.x + (lines[caretLine]!.length === 1 ? 10 : 0),
        y: 108 + caretLine * 20, height: 16 } }),
  } as unknown as DocxWorkerCommandEngine;
  const at = { story: 'body', paraId: 'p', offset: lineStart + 1 };
  const alphabet = 'abcdeAV ';
  const glyphSet = (kern: boolean) => {
    const set = { key: kern ? 'kern' : 'set', fontId: 1, alphabet, horizontal: 0.01,
      glyphIds: [...alphabet].map((char) => char.charCodeAt(0)), advances: [...alphabet].map(() => 1000),
      kerning: kern ? [[5, 6, -200] as const] : [], unsafePairs: [], outlines: {} };
    return { key: set.key, ink: { upem: 1000, xMin: 0, yMin: -200, yMax: 700, overhang: 0 }, glyphs: { set, horizontal: 0.01, advance: (units: number) => units * 0.01,
      chars: new Map([...alphabet].map((char, index) => [char, index])),
      kerning: new Map(set.kerning.map(([left, right, units]) => [left * alphabet.length + right, units])),
      unsafe: new Set<number>() } };
  };
  const sets: unknown[][] = [];
  const glyphs: EchoGlyphProvider = { pixelRatio: 1,
    glyphSet(fontId, size, kern) {
      sets.push([fontId, size, kern]);
      return glyphSet(kern);
    },
    glyphInk: () => ({ upem: 1000, xMin: 0, xMax: 500, yMin: -200, yMax: 700 }),
  };
  return { engine, selection: { anchor: at, focus: at }, layout: { frame, layoutRevision: 2 }, glyphs, sets };
}

test('glyph line reproduces the engine placement of a wrapped continuation line', () => {
  const source = glyphLine();
  const context = projectCommandEchoContext(source);
  expect(source.sets).toEqual([[1, Math.fround(960 / 72), false]]);
  expect(context?.glyphs).toEqual({ glyphSet: 'set', fontId: 1, size: 13.333, color: '#000000', x: 96, y: 140,
    next: '', right: 720, previousSlack: 594, device: { x: 95, y: 129, width: 626, height: 15 }, pixelRatio: 1 });
  expect(context?.damage[0]?.rect).toEqual({ x: 95, y: 129, width: 626, height: 15 });
  expect(context?.damage[0]?.omittedPrimitiveIds).toEqual(['4', '5']);
  const first = projectCommandEchoContext(glyphLine({ caretLine: 0 }));
  expect(first?.glyphs).toMatchObject({ next: 'c', nextWord: 20, y: 120 });
  expect(first?.glyphs?.previousSlack).toBeUndefined();
});

test('glyph line keeps only the glyph sets that reproduce the line spacing and kerning', () => {
  expect(projectCommandEchoContext(glyphLine({ shiftX: 0.001 }))).toBeUndefined();
  const { glyphs: _, ...withoutGlyphs } = glyphLine({ shiftX: 0.001 });
  expect(projectCommandEchoContext(withoutGlyphs)?.text).toBe('cd');
  expect(projectCommandEchoContext(glyphLine({ advance: 11 }))).toBeUndefined();
  expect(projectCommandEchoContext(glyphLine({ lines: ['ab ', 'AVc'], kerned: true, kerning: 2 }))?.glyphs?.glyphSet)
    .toBe('kern');
  expect(projectCommandEchoContext(glyphLine({ lines: ['ab ', 'AVc'], kerned: true }))).toBeUndefined();
  expect(projectCommandEchoContext(glyphLine({ lines: ['ab ', 'AVc'], kerning: 2 }))).toBeUndefined();
});

test('glyph line declines lines whose origin or stretch moves with their width', () => {
  expect(projectCommandEchoContext(glyphLine({ x: 146 }))).toBeUndefined();
  expect(projectCommandEchoContext(glyphLine({ right: 116 }))).toBeUndefined();
  expect(projectCommandEchoContext(glyphLine({ right: 117 }))?.glyphs).toBeDefined();
});

test('glyph line declines decorations and backgrounds bound to its text', () => {
  const underline = { kind: 'decoration' as const, deco: 'underline' as const, x: 96, y: 142, w: 20, h: 1, color: '#f00',
    docStart: 4, docEnd: 6 };
  expect(projectCommandEchoContext(glyphLine({ later: [underline] }))).toBeUndefined();
  const highlight = { kind: 'rect' as const, x: 96, y: 128, w: 10, h: 16, fill: '#ff0', docStart: 4, docEnd: 5 };
  expect(projectCommandEchoContext(glyphLine({ later: [highlight] }))).toBeUndefined();
});

test('glyph line declines when later paint reaches its band', () => {
  expect(projectCommandEchoContext(glyphLine({ later: [{ kind: 'rect', x: 300, y: 138, w: 20, h: 2, fill: '#f00' }] })))
    .toBeUndefined();
  expect(projectCommandEchoContext(glyphLine({ later: [{ kind: 'rect', x: 300, y: 200, w: 20, h: 2, fill: '#f00' }] }))
    ?.glyphs).toBeDefined();
});

test('echo uses the resolved right edge of a style with a 1440 twip right indent', () => {
  const source = glyphLine({ lines: ['ab ', 'cd'] });
  for (const run of source.layout.frame.pages[0]!.page.primitives) {
    if (run.kind === 'glyphRun') Object.assign(run, { echo: { x: 96, right: 624, kern: false, primaryFont: true } });
  }
  expect(projectCommandEchoContext(source)?.glyphs?.right).toBe(624);
});

test('echo declines a single character with expanded tracking', () => {
  const source = glyphLine({ lines: ['c'] });
  Object.assign(source.engine, { paragraphSegments: () => [[{ kind: 'text', text: 'c',
    attributes: { characterSpacing: { spacing: 40 } } }]] });
  for (const run of source.layout.frame.pages[0]!.page.primitives) {
    if (run.kind === 'glyphRun') Object.assign(run, { echo: { x: 96, right: 720, kern: false, primaryFont: true,
      letterSpacing: 40 / 15 } });
  }
  expect(projectCommandEchoContext(source)).toBeUndefined();
});

test('echo requires resolved geometry and default spacing on every run', () => {
  for (const change of [{ letterSpacing: -1 }, { position: 1 }, { scale: 80 }]) {
    const source = glyphLine();
    const run = source.layout.frame.pages[0]!.page.primitives[4] as GlyphRunPrimitive;
    Object.assign(run.echo!, change);
    expect(projectCommandEchoContext(source)).toBeUndefined();
  }
  const source = glyphLine();
  delete (source.layout.frame.pages[0]!.page.primitives[4] as GlyphRunPrimitive).echo;
  expect(projectCommandEchoContext(source)).toBeUndefined();
  const thresholds = glyphLine({ kerning: 40 });
  const run = thresholds.layout.frame.pages[0]!.page.primitives[4] as GlyphRunPrimitive;
  run.echo!.kerning = 30;
  expect(projectCommandEchoContext(thresholds)).toBeUndefined();
});

test('echo declines substitution starters on another line of the paragraph', () => {
  const source = glyphLine({ lines: ['ab ', 'de'], right: 130 });
  const build = source.glyphs.glyphSet;
  source.glyphs.glyphSet = (font, size, kern) => {
    const set = build(font, size, kern)!;
    return { ...set, glyphs: { ...set.glyphs, unsafe: new Set([1]) } };
  };
  expect(projectCommandEchoContext(source)).toBeUndefined();
});

test('echo declines fallback letters when a digits-only primary will resolve future digits', async () => {
  const { readFileSync } = await import('node:fs');
  const { preloadEditWasm } = await import('./wasm/index');
  const { createResidentEngineSession } = await import('./residentEngineSession');
  await preloadEditWasm(new Uint8Array(readFileSync('packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm')));
  const bytes = new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'));
  const primary = new Uint8Array(bytes.length + 40);
  primary.set(bytes);
  const view = new DataView(primary.buffer);
  const cmap = bytes.length;
  view.setUint16(cmap + 2, 1);
  view.setUint16(cmap + 4, 3);
  view.setUint16(cmap + 6, 10);
  view.setUint32(cmap + 8, 12);
  view.setUint16(cmap + 12, 12);
  view.setUint32(cmap + 16, 28);
  view.setUint32(cmap + 24, 1);
  view.setUint32(cmap + 28, 48);
  view.setUint32(cmap + 32, 57);
  view.setUint32(cmap + 36, 36);
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const entry = 12 + index * 16;
    if (view.getUint32(entry) === 0x636d6170) {
      view.setUint32(entry + 8, cmap);
      view.setUint32(entry + 12, 40);
    }
  }
  const engine = await createResidentEngineSession();
  try {
    const digits = engine.registerFont(primary);
    const fallback = engine.registerFont(bytes);
    expect(JSON.parse(engine.echoGlyphSetJson(digits, 13.333333, false, 'abc1')).glyphIds.slice(0, 3)).toEqual([0, 0, 0]);
    const shaped = JSON.parse(engine.echoGlyphSetJson(fallback, Math.fround(960 / 72), false, 'abc1'));
    expect(JSON.parse(engine.echoGlyphSetJson(digits, 13.333333, false, '1')).glyphIds[0]).not.toBe(shaped.glyphIds[3]);
    const glyphs = indexEchoGlyphSet({ key: 'fallback', fontId: fallback, alphabet: 'abc1', ...shaped, outlines: {} });
    const placement = placeEchoGlyphs(glyphs, 'abc', '', 96)!;
    const source = glyphLine({ lines: ['abc'] });
    for (const [index, run] of source.layout.frame.pages[0]!.page.primitives.entries()) {
      if (run.kind !== 'glyphRun') continue;
      run.fontId = fallback;
      run.glyphs[0] = { ...run.glyphs[0]!, id: placement.glyphIds[index]!, x: placement.xs[index]!,
        advance: placement.advances[index]! };
      Object.assign(run, { echo: { x: 96, right: 720, kern: false, primaryFont: false } });
    }
    Object.assign(source.engine, { residentCaretSnapshot: () => ({ frameEpoch: 3,
      caretRect: { pageId: '1', pageIndex: 0, x: placement.xs[1], y: 108, height: 16 } }) });
    source.glyphs.glyphSet = () => ({ key: 'fallback', glyphs,
      ink: { upem: 2048, xMin: 0, yMin: -400, yMax: 1500, overhang: 0 } });
    expect(projectCommandEchoContext(source)).toBeUndefined();
  } finally { engine.destroy(); }
}, 20000);

test('echo projection allocates no encoders or byte arrays per character', () => {
  const source = glyphLine({ lines: ['a'.repeat(40)], caretLine: 0 });
  const Encoder = globalThis.TextEncoder;
  const encode = Encoder.prototype.encode;
  let constructors = 0;
  let characters = 0;
  globalThis.TextEncoder = class extends Encoder { constructor() { super(); constructors += 1; } };
  Encoder.prototype.encode = function (text = '') {
    if ([...text].length === 1) characters += 1;
    return encode.call(this, text);
  };
  try {
    expect(projectCommandEchoContext(source)?.text).toBe('a'.repeat(40));
  } finally {
    Encoder.prototype.encode = encode;
    globalThis.TextEncoder = Encoder;
  }
  expect(constructors).toBe(0);
  expect(characters).toBe(0);
});

test('echo admits ordinary letters when contextual lookups require combining marks outside the alphabet', async () => {
  const { readFileSync } = await import('node:fs');
  const { preloadEditWasm } = await import('./wasm/index');
  const { createResidentEngineSession } = await import('./residentEngineSession');
  const { DOCX_ECHO_ALPHABET } = await import('./residentEchoGlyphs');
  await preloadEditWasm(new Uint8Array(readFileSync('packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm')));
  const engine = await createResidentEngineSession();
  try {
    const font = engine.registerFont(new Uint8Array(readFileSync('packages/fonts/assets/Carlito-Regular.ttf')));
    const set = JSON.parse(engine.echoGlyphSetJson(font, Math.fround(960 / 72), false, DOCX_ECHO_ALPHABET));
    expect(set.glyphIds[DOCX_ECHO_ALPHABET.indexOf('i')]).toBeGreaterThan(0);
    expect(set.glyphIds[DOCX_ECHO_ALPHABET.indexOf('j')]).toBeGreaterThan(0);
  } finally { engine.destroy(); }
}, 20000);
