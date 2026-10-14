import { expect, test } from 'bun:test';
import type { DisplayPage, GlyphRunPrimitive } from '../layout/render/displayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { ResidentDocxCommandEngine } from './residentCommandProtocol';
import { projectCommandEchoContext } from './residentCommandViewProjection';

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
  } as unknown as ResidentDocxCommandEngine;
  const at = { story: 'body', paraId: 'p', offset: 3 };
  return { engine, selection: { anchor: at, focus: at }, layout: { frame, layoutRevision: 2 }, run };
}

test('d2c_projection_maps_byte_clusters_to_utf16_graphemes', () => {
  const source = input();
  const context = projectCommandEchoContext(source);
  expect(context?.boundaries.map(({ utf16, grapheme, caret }) => [utf16, grapheme, caret.x]))
    .toEqual([[0, 0, 0], [1, 1, 10], [3, 2, 20], [4, 3, 30]]);
  expect(context?.selection).toEqual({ anchorGrapheme: 2, focusGrapheme: 2 });
  expect(context?.damage[0]?.primitiveIds).toEqual(['10', '12']);
  expect(context?.serializedBytes).toBe(new TextEncoder().encode(JSON.stringify(context)).byteLength);
  expect(projectCommandEchoContext({ ...source, selection: null })).toBeUndefined();
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
