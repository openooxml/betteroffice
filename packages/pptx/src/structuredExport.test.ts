import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { PptxExportParagraph, PptxExportShape, PptxStructuredContent } from './index';
import {
  exportPptxMarkdown,
  exportPptxStructured,
  initWasm,
  openPresentation,
  PptxExportError,
  renderPptxMarkdown,
} from './index';

const root = resolve(import.meta.dir, '../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

function paragraphs(shapes: PptxExportShape[]): PptxExportParagraph[] {
  return shapes.flatMap((shape) => [
    ...shape.stories.flatMap((story) => story.paragraphs),
    ...(shape.table?.rows ?? []).flatMap((row) =>
      row.cells.flatMap((cell) => cell.story?.paragraphs ?? [])
    ),
    ...paragraphs(shape.children),
  ]);
}

function content(result: ReturnType<ReturnType<typeof openPresentation>['exportStructured']>) {
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

describe('structured export', () => {
  test('a session export reads committed state and changes nothing', () => {
    const deck = openPresentation(fixture, { clientId: 9401 });
    try {
      let updates = 0;
      const unsubscribe = deck.onUpdate(() => {
        updates += 1;
      });
      const version = deck.version();
      const state = deck.encodeStateAsUpdate();
      const read = content(deck.exportStructured());
      expect(read.version).toBe(version);
      expect(read.content.schemaVersion).toBe(1);
      expect(read.content.anchorScope).toBe('session');
      expect(read.content.readingOrder).toBe('shapeTree');
      expect(read.content.slides.map((slide) => slide.index)).toEqual([0, 1, 2]);

      const stories = deck.readContent();
      if (!stories.ok) throw new Error(stories.failure.message);
      const exported = new Map(
        read.content.slides
          .flatMap((slide) => paragraphs(slide.shapes))
          .map((paragraph) => [paragraph.paragraphId, paragraph.anchor])
      );
      let compared = 0;
      for (const story of stories.stories) {
        for (const paragraph of story.paragraphs) {
          const anchor = exported.get(paragraph.paragraphId);
          if (!anchor) continue;
          expect(anchor).toEqual({
            kind: 'range',
            slideId: story.slideId,
            shapeId: story.shapeId,
            storyId: story.storyId,
            start: paragraph.start,
            end: paragraph.end,
          });
          compared += 1;
        }
      }
      expect(compared).toBeGreaterThan(20);

      const markdown = deck.exportMarkdown({ includeNotes: true });
      if (!markdown.ok) throw new Error(markdown.failure.message);
      expect(markdown.version).toBe(version);
      expect(markdown.content.markdown.match(/<!-- pptx-export:\d+ -->/g)?.length).toBe(
        markdown.content.anchors.length
      );

      const refused = deck.exportStructured({ maxBytes: 8 });
      expect(refused).toEqual({
        ok: false,
        version,
        failure: {
          code: 'invalid-options',
          target: null,
          message: 'maxBytes must be at least 1024',
        },
      });
      expect(() => deck.exportStructured({ maxBytes: Number.NaN })).toThrow(RangeError);
      expect(() => deck.exportStructured({ pages: true } as never)).toThrow();

      expect(deck.version()).toBe(version);
      expect(deck.encodeStateAsUpdate()).toEqual(state);
      expect(deck.canUndo()).toBe(false);
      expect(updates).toBe(0);
      unsubscribe();
    } finally {
      deck.dispose();
    }
  });

  test('a range anchor is a batch target', () => {
    const deck = openPresentation(fixture, { clientId: 9402 });
    try {
      const read = content(deck.exportStructured());
      const anchor = paragraphs(read.content.slides[0].shapes)
        .flatMap((paragraph) => paragraph.runs)
        .find((run) => run.kind === 'text')?.anchor;
      if (anchor?.kind !== 'range') throw new Error('no text run');
      const applied = deck.applyEdits({
        expectVersion: read.version,
        steps: [{ op: 'replaceText', target: anchor, text: 'Replaced' }],
      });
      expect(applied.ok).toBe(true);
      const stories = deck.readContent({ slideIds: [anchor.slideId] });
      if (!stories.ok) throw new Error(stories.failure.message);
      const story = stories.stories.find((candidate) => candidate.storyId === anchor.storyId);
      expect(story?.text.slice(anchor.start, anchor.start + 8)).toBe('Replaced');
    } finally {
      deck.dispose();
    }
  });

  test('bytes exports are snapshots that render the same Markdown', async () => {
    const structured: PptxStructuredContent = await exportPptxStructured(fixture);
    expect(structured.anchorScope).toBe('snapshot');
    expect(structured.slides[0]?.provenance?.part).toBe('ppt/slides/slide1.xml');
    expect(structured.slides[0]?.provenance?.partSha256).toMatch(/^[0-9a-f]{64}$/);
    expect(JSON.stringify(await exportPptxStructured(fixture))).toBe(JSON.stringify(structured));

    const rendered = await renderPptxMarkdown(structured);
    expect(rendered).toEqual(await exportPptxMarkdown(fixture));
    expect(rendered.markdown.startsWith('<!-- pptx-export:0 -->\n## Slide 1')).toBe(true);

    const truncated = await exportPptxStructured(fixture, { maxBytes: 4_096 });
    expect(truncated.truncated).toBe(true);
    expect(new TextEncoder().encode(JSON.stringify(truncated)).length).toBeLessThanOrEqual(4_096);
    expect(truncated.diagnostics[truncated.diagnostics.length - 1]?.code).toBe('truncated');

    const refusal = await exportPptxStructured(fixture, { maxBlocks: 0 }).catch((error) => error);
    expect(refusal).toBeInstanceOf(PptxExportError);
    expect((refusal as PptxExportError).failure.code).toBe('invalid-options');
    const unreadable = await exportPptxStructured(new Uint8Array([1, 2, 3])).catch(
      (error) => error
    );
    expect(unreadable).toBeInstanceOf(Error);
    expect(unreadable).not.toBeInstanceOf(PptxExportError);
    const invalid = await renderPptxMarkdown({ ...structured, schemaVersion: 2 as 1 }).catch(
      (error) => error
    );
    expect(invalid).toBeInstanceOf(Error);
  });
});
