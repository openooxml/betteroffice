import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type {
  PresentationHandle,
  ShapePrimitive,
  StorySnapshot,
  TextBoxPrimitive,
} from '../index';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
let handle: PresentationHandle;
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  handle = openPresentation(pptx, {
    clientId: 9001,
    fonts: [{ family: 'Liberation Sans', bytes: font }],
  });
});

afterAll(() => handle.dispose());

describe('PPTX wasm boundary', () => {
  test('opens shared updates without parsing the file bytes', () => {
    const source = openPresentation(fixture, { clientId: 9002 });
    const seed = source.encodeStateAsUpdate();
    const left = openPresentation(Uint8Array.of(0xff), {
      clientId: 9003,
      initialUpdate: seed,
    });
    const right = openPresentation(Uint8Array.of(0xff), {
      clientId: 9004,
      initialUpdate: seed,
    });

    expect(left.snapshot()).toEqual(source.snapshot());
    expect([...left.encodeStateAsUpdate()]).toEqual([...seed]);
    expect([...right.encodeStateAsUpdate()]).toEqual([...seed]);
    expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);

    source.dispose();
    left.dispose();
    right.dispose();
  });

  test('opens, edits, reflows, hit-tests, and observes a local update', () => {
    const snapshot = handle.snapshot();
    expect(snapshot.slides.length).toBe(3);
    const story = firstStory(snapshot.slides.flatMap((slide) => slide.shapes));
    const insertion = story.length - 1;

    const events: Array<{ origin: string; update: Uint8Array }> = [];
    const unsubscribe = handle.onUpdate((update, origin) => events.push({ origin, update }));
    const receipt = handle.insertText(story.id, insertion, ' edited', {
      bold: true,
      fontSizePt: 28,
      color: '#325ee6',
    });
    expect(receipt.storyId).toBe(story.id);
    expect(handle.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('edited'))
    )).toBe(true);

    const frame = handle.layoutSlide(0);
    const textBox = frame.primitives.find(
      (primitive): primitive is TextBoxPrimitive =>
        primitive.kind === 'textBox' && primitive.storyId === story.id
    );
    expect(textBox?.lines.length).toBeGreaterThan(0);
    const line = textBox!.lines[0];
    expect(handle.hitTest(line.x, line.y + line.height / 2)?.kind).toBe('text');

    expect(events[0]?.origin).toBe('local');
    expect(events[0]?.update.length).toBeGreaterThan(0);
    expect(handle.canUndo()).toBe(true);
    expect(handle.undo().applied).toBe(true);
    expect(handle.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('edited'))
    )).toBe(false);
    unsubscribe();
  });

  test('inserts and styles preset shapes with undo and redo', () => {
    const slide = handle.snapshot().slides[0];
    const receipt = handle.addShape(slide.id, {
      name: 'Styled rounded rectangle',
      geometry: 'roundRect',
      rect: { x: 900_000, y: 1_000_000, width: 3_100_000, height: 1_400_000 },
      fill: '#D9EAF7',
    });
    expect(shapeSnapshot(receipt.shapeId).geometry).toBe('roundRect');
    expect(handle.undo().snapshot.slides[0].shapes.some(
      (shape) => shape.id === receipt.shapeId
    )).toBe(false);
    expect(handle.redo().snapshot.slides[0].shapes.some(
      (shape) => shape.id === receipt.shapeId
    )).toBe(true);

    handle.setShapeFill(slide.id, receipt.shapeId, '#3367D6');
    expect(shapeSnapshot(receipt.shapeId).fill?.color?.rgb).toBe('3367D6');
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).fill?.color?.rgb).toBe('D9EAF7');
    expect(handle.redo().applied).toBe(true);

    handle.setShapeStroke(slide.id, receipt.shapeId, {
      color: '#EA4335',
      widthPt: 3,
    });
    expect(shapeSnapshot(receipt.shapeId).outline?.width).toBe(38_100);
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).outline).toBeNull();
    expect(handle.redo().applied).toBe(true);

    handle.setShapeAdjust(slide.id, receipt.shapeId, { adj: 0.32 });
    expect(shapeSnapshot(receipt.shapeId).adjustValues.adj).toBe(0.32);
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).adjustValues.adj).toBeCloseTo(0.16667);
    expect(handle.redo().applied).toBe(true);

    const primitive = handle.layoutSlide(0).primitives.find(
      (candidate): candidate is ShapePrimitive =>
        candidate.kind === 'shape' && candidate.shapeId === receipt.shapeId
    );
    expect(primitive?.fill).toEqual({ kind: 'solid', color: '#3367D6' });
    expect(primitive?.stroke?.color).toBe('#EA4335');
    expect(primitive?.adjustValues?.adj).toBeCloseTo(0.32);

    handle.setShapeFill(slide.id, receipt.shapeId, null);
    handle.setShapeStroke(slide.id, receipt.shapeId, {});
    const cleared = handle.layoutSlide(0).primitives.find(
      (candidate): candidate is ShapePrimitive =>
        candidate.kind === 'shape' && candidate.shapeId === receipt.shapeId
    );
    expect(cleared?.fill).toBeUndefined();
    expect(cleared?.stroke).toBeUndefined();
  });

  test('a session opened from an update saves when the source file is attached', () => {
    const seeded = openPresentation(fixture, { clientId: 9007 });
    const seed = seeded.encodeStateAsUpdate();

    const attached = openPresentation(fixture, { clientId: 9008, initialUpdate: seed });
    const slide = attached.snapshot().slides[0];
    attached.moveShape(slide.id, slide.shapes[0].id, 777_000, 888_000);
    const reopened = openPresentation(attached.save(), { clientId: 9009 });
    const moved = reopened.snapshot().slides[0].shapes[0];
    expect([moved.x, moved.y]).toEqual([777_000, 888_000]);

    const bare = openPresentation(Uint8Array.of(0xff), { clientId: 9010, initialUpdate: seed });
    expect(() => bare.save()).toThrow(/source file bytes/);

    seeded.dispose();
    attached.dispose();
    reopened.dispose();
    bare.dispose();
  });

  test('edits survive a save and reopen', () => {
    const source = openPresentation(fixture, { clientId: 9005 });
    const slide = source.snapshot().slides[0];
    const shape = slide.shapes.find((candidate) => candidate.sourceId !== 0)!;
    source.moveShape(slide.id, shape.id, 1_234_000, 2_345_000);
    const story = firstStory(slide.shapes);
    source.insertText(story.id, 0, 'Saved: ');

    const reopened = openPresentation(source.save(), { clientId: 9006 });
    const snapshot = reopened.snapshot();
    const moved = snapshot.slides[0].shapes.find(
      (candidate) => candidate.sourceId === shape.sourceId
    );
    expect([moved?.x, moved?.y]).toEqual([1_234_000, 2_345_000]);
    const text = snapshot.slides
      .flatMap((candidate) => candidate.shapes)
      .flatMap((candidate) => candidate.textStories)
      .find((candidate) => candidate.id === story.id);
    expect(text?.paragraphs[0]?.runs[0]?.text.startsWith('Saved: ')).toBe(true);

    source.dispose();
    reopened.dispose();
  });
});

function shapeSnapshot(shapeId: string) {
  const shape = handle.snapshot().slides[0].shapes.find((candidate) => candidate.id === shapeId);
  if (!shape) throw new Error(`shape ${shapeId} was not found`);
  return shape;
}

function firstStory(shapes: Array<{ textStories: StorySnapshot[]; children: unknown[] }>): StorySnapshot {
  for (const shape of shapes) {
    if (shape.textStories[0]) return shape.textStories[0];
  }
  throw new Error('fixture has no text story');
}
