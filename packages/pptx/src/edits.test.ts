import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type {
  PptxEditRequest,
  PptxReadResult,
  PptxStoryTarget,
  PptxStoryText,
  PptxTextTarget,
  PresentationHandle,
} from './index';
import { initWasm, openPresentation } from './index';

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

function withDeck(clientId: number, run: (deck: PresentationHandle) => void): void {
  const deck = openPresentation(fixture, { clientId });
  try {
    run(deck);
  } finally {
    deck.dispose();
  }
}

function read(deck: PresentationHandle): Extract<PptxReadResult, { ok: true }> {
  const result = deck.readContent();
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

const firstStory = (result: Extract<PptxReadResult, { ok: true }>): PptxStoryText =>
  result.stories.find((story) => /^\w+ /.test(story.text))!;

const within = (story: PptxStoryText): PptxStoryTarget => ({
  slideId: story.slideId,
  shapeId: story.shapeId,
  storyId: story.storyId,
});

const range = (story: PptxStoryText, start: number, end: number): PptxTextTarget => ({
  kind: 'range',
  ...within(story),
  start,
  end,
});

describe('version-checked edit batches', () => {
  test('validate reserves nothing, and apply commits one update and one undo step', () => {
    withDeck(9301, (deck) => {
      const current = read(deck);
      expect(current.version).toBe(deck.version());
      const story = firstStory(current);
      const updates: Uint8Array[] = [];
      deck.onUpdate((update) => updates.push(update));
      const request: PptxEditRequest = {
        expectVersion: current.version,
        steps: [{ op: 'insertText', target: range(story, 0, 0), at: 'start', text: 'Draft: ' }],
      };
      expect(deck.validateEdits(request)).toMatchObject({
        ok: true,
        baseVersion: current.version,
        wouldApply: true,
        previews: [{ stepIndex: 0, wouldChange: true, target: range(story, 0, 0) }],
      });
      expect(deck.version()).toBe(current.version);
      expect(updates).toHaveLength(0);

      const result = deck.applyEdits(request);
      if (!result.ok) throw new Error(result.failure.message);
      expect(result).toMatchObject({
        applied: true,
        source: 'host',
        baseVersion: current.version,
        version: deck.version(),
        changedStories: [story.storyId],
        changedSlides: [story.slideId],
        receipts: [{ stepIndex: 0, changed: true, target: range(story, 0, 7) }],
      });
      expect(updates).toHaveLength(1);
      expect(read(deck).stories.find((candidate) => candidate.storyId === story.storyId)?.text)
        .toBe(`Draft: ${story.text}`);

      expect(deck.applyEdits(request)).toMatchObject({
        ok: false,
        version: deck.version(),
        failure: { code: 'stale-version' },
      });
      expect(updates).toHaveLength(1);
      deck.undo();
      expect(deck.canUndo()).toBe(false);
      expect(read(deck).stories.find((candidate) => candidate.storyId === story.storyId)?.text)
        .toBe(story.text);
    });
  });

  test('refusals are data, malformed requests throw, and reads refuse missing slides', () => {
    withDeck(9302, (deck) => {
      const current = read(deck);
      const story = firstStory(current);
      const before = deck.snapshot();
      const refusal = deck.applyEdits({
        expectVersion: current.version,
        steps: [
          { op: 'insertText', target: range(story, 0, 0), at: 'start', text: 'x' },
          {
            op: 'deleteText',
            target: { kind: 'search', within: within(story), text: 'no such text anywhere' },
          },
        ],
      });
      expect(refusal).toMatchObject({
        ok: false,
        version: current.version,
        failure: { code: 'missing-target', stepIndex: 1, target: { kind: 'search' } },
      });
      expect(deck.snapshot()).toEqual(before);
      expect(deck.canUndo()).toBe(false);
      expect(() =>
        deck.applyEdits({
          expectVersion: current.version,
          steps: [{ op: 'explode' }],
        } as unknown as PptxEditRequest)
      ).toThrow();
      expect(() => deck.validateEdits({ steps: [] } as unknown as PptxEditRequest)).toThrow();
      expect(deck.readContent({ slideIds: ['slide:missing'] })).toMatchObject({
        ok: false,
        failure: { code: 'missing-target', target: { kind: 'slide', slideId: 'slide:missing' } },
      });
      const found = deck.findText({ text: 'e', limit: 2 });
      expect(found).toMatchObject({ ok: true, version: current.version, truncated: true });
      if (!found.ok) return;
      expect(found.matches).toHaveLength(2);
      expect(found.matches[0].text).toBe('e');
    });
  });

  test('non-finite numbers throw before they could reach the engine as null', () => {
    withDeck(9305, (deck) => {
      const slideId = deck.snapshot().slides[0].id;
      const { shapeId } = deck.addShape(slideId, {
        name: 'Outlined',
        geometry: 'rect',
        rect: { x: 0, y: 0, width: 914_400, height: 914_400 },
      });
      deck.setShapeStroke(slideId, shapeId, { color: '#112233', widthPt: 2 });
      const version = deck.version();
      const outline = () =>
        deck.snapshot().slides[0].shapes.find((shape) => shape.id === shapeId)?.outline;
      const before = outline();
      for (const widthPt of [Number.NaN, Number.POSITIVE_INFINITY]) {
        const request: PptxEditRequest = {
          expectVersion: version,
          steps: [{ op: 'setShapeStroke', target: { slideId, shapeId }, stroke: { widthPt } }],
        };
        expect(() => deck.validateEdits(request)).toThrow(RangeError);
        expect(() => deck.applyEdits(request)).toThrow(RangeError);
      }
      expect(() => deck.findText({ text: 'e', limit: Number.NaN })).toThrow(RangeError);
      expect(outline()).toEqual(before);
      expect(before?.width).toBe(2 * 12_700);
      expect(deck.version()).toBe(version);
    });
  });

  test('untracked agent batches save, reopen, and never carry versions across sessions', () => {
    withDeck(9303, (deck) => {
      const current = read(deck);
      const story = firstStory(current);
      const word = story.text.split(' ')[0];
      const slide = current.slides.find((candidate) => candidate.id === story.slideId)!;
      const shape = slide.shapes.find((candidate) => candidate.id === story.shapeId)!;
      const rect = { x: shape.x, y: shape.y, width: shape.width, height: shape.height };
      const result = deck.applyEdits({
        expectVersion: current.version,
        source: 'agent',
        history: 'none',
        steps: [
          {
            op: 'replaceText',
            target: { kind: 'search', within: within(story), text: word },
            text: 'Revised',
            expect: { text: word },
          },
          {
            op: 'setShapeRect',
            target: { slideId: slide.id, shapeId: shape.id },
            rect: { ...rect, x: rect.x + 12_700 },
            expect: { rect },
          },
        ],
      });
      expect(result).toMatchObject({ ok: true, applied: true, source: 'agent' });
      expect(deck.canUndo()).toBe(false);
      const reopened = openPresentation(deck.save(), { clientId: 9304 });
      try {
        const again = read(reopened);
        expect(again.stories.find((candidate) => candidate.storyId === story.storyId)?.text)
          .toStartWith('Revised');
        expect(again.slides[0].shapes.find((candidate) => candidate.id === shape.id)?.x)
          .toBe(rect.x + 12_700);
        expect(reopened.version()).not.toBe(deck.version());
        expect(
          reopened.applyEdits({
            expectVersion: deck.version(),
            steps: [{ op: 'deleteText', target: range(story, 0, 1) }],
          })
        ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      } finally {
        reopened.dispose();
      }
    });
  });
});
