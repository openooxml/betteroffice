import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type {
  PresentationHandle,
  ShapePrimitive,
  SlidePrimitive,
  StorySnapshot,
  TextBoxPrimitive,
} from '../index';
import { initWasm, openPresentation, paintSlide } from '../index';

const root = resolve(import.meta.dir, '../../../..');
let handle: PresentationHandle;
let fixture: Uint8Array;
let fontBytes: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  fontBytes = font;
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

  // the wasm event arrives as `[origin, ...update]`, so a listener that reaches
  // for `.buffer` must not find the origin tag riding along — and the shape must
  // not depend on how many other listeners happen to be subscribed.
  test('delivers exact update buffers regardless of subscriber count', () => {
    const seed = openPresentation(fixture, { clientId: 9201 });
    let update: Uint8Array;
    try {
      update = seed.encodeStateAsUpdate();
    } finally {
      seed.dispose();
    }
    for (const extraSubscriber of [false, true]) {
      const source = openPresentation(Uint8Array.of(0xff), {
        clientId: extraSubscriber ? 9202 : 9203,
        initialUpdate: update,
      });
      const peer = openPresentation(Uint8Array.of(0xff), {
        clientId: extraSubscriber ? 9204 : 9205,
        initialUpdate: update,
      });
      const received: Uint8Array[] = [];
      try {
        source.onUpdate((bytes, origin) => {
          if (origin === 'local') received.push(bytes);
        });
        if (extraSubscriber) source.onUpdate(() => {});

        const story = firstStory(source.snapshot().slides.flatMap((slide) => slide.shapes));
        source.insertText(story.id, story.length - 1, ' exact');
        expect(received).toHaveLength(1);
        expect(received[0].byteOffset).toBe(0);
        expect(received[0].buffer.byteLength).toBe(received[0].byteLength);

        peer.applyUpdate(new Uint8Array(received[0].buffer));
        expect(peer.story(story.id)).toEqual(source.story(story.id));
      } finally {
        source.dispose();
        peer.dispose();
      }
    }
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

  test('aligns paragraphs, reflows the text box, and undoes in one step', () => {
    const snapshot = handle.snapshot();
    const story = firstStory(snapshot.slides.flatMap((slide) => slide.shapes));
    const laidOut = () =>
      handle.layoutSlide(0).primitives.find(
        (primitive): primitive is TextBoxPrimitive =>
          primitive.kind === 'textBox' && primitive.storyId === story.id
      );
    const storedBefore = handle.story(story.id).paragraphs[0].alignment;
    const renderedBefore = laidOut()?.paragraphs[0]?.align;

    handle.setParagraphAlignment(story.id, 0, 0, 'ctr');
    expect(handle.story(story.id).paragraphs[0].alignment).toBe('ctr');
    expect(laidOut()?.paragraphs[0]?.align).toBe('center');

    expect(handle.undo().applied).toBe(true);
    expect(handle.story(story.id).paragraphs[0].alignment).toBe(storedBefore);
    expect(laidOut()?.paragraphs[0]?.align).toBe(renderedBefore);

    expect(() =>
      handle.setParagraphAlignment(story.id, 0, 0, 'middle' as never)
    ).toThrow();
  });

  test('sets a shape rectangle in one update and one undo step', () => {
    const source = openPresentation(fixture, { clientId: 9011 });
    const slide = source.snapshot().slides[0];
    const shape = slide.shapes[0];
    const before = {
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height,
    };
    const after = {
      x: before.x + 120_000,
      y: before.y + 80_000,
      width: before.width + 300_000,
      height: before.height + 200_000,
    };
    const updates: Uint8Array[] = [];
    source.onUpdate((update, origin) => {
      if (origin === 'local') updates.push(update);
    });

    expect(source.setShapeRect(slide.id, shape.id, after).after).toEqual(after);
    expect(updates).toHaveLength(1);
    expect(shapeSnapshotFrom(source, shape.id)).toMatchObject(after);
    expect(source.undo().applied).toBe(true);
    expect(shapeSnapshotFrom(source, shape.id)).toMatchObject(before);
    expect(source.canUndo()).toBe(false);
    source.dispose();
  });

  test('paints the non-justified demo deck with one call per text run', async () => {
    const source = openPresentation(fixture, {
      clientId: 9013,
      fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    });
    const calls: Array<{ text: string; x: number; y: number }> = [];
    const expected: Array<{ text: string; x: number; y: number }> = [];
    const ctx = new Proxy(
      {
        fillText: (text: string, x: number, y: number) => calls.push({ text, x, y }),
      } as Record<string, unknown>,
      {
        get(target, property) {
          if (property in target) return target[property as string];
          return () => undefined;
        },
        set(target, property, value) {
          target[property as string] = value;
          return true;
        },
      }
    ) as unknown as CanvasRenderingContext2D;
    try {
      for (let slideIndex = 0; slideIndex < source.snapshot().slides.length; slideIndex += 1) {
        const frame = source.layoutSlide(slideIndex);
        expected.push(...oneCallPerTextRun(frame.primitives));
        await paintSlide(ctx, frame);
      }

      expect(expected).toHaveLength(62);
      expect(calls).toHaveLength(expected.length);
      expect(calls).toEqual(expected);
    } finally {
      source.dispose();
    }
  });

  test('paints engine-produced justified word starts at caret positions', async () => {
    const source = openPresentation(fixture, {
      clientId: 9012,
      fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    });
    try {
      const shape = source.snapshot().slides[0].shapes.find(
        (candidate) => candidate.name === 'Subtitle'
      );
      const story = shape?.textStories[0];
      if (!story) throw new Error('subtitle story is missing');
      source.setParagraphAlignment(story.id, 0, story.length, 'just');
      const frame = source.layoutSlide(0);
      const textBox = frame.primitives.find(
        (primitive): primitive is TextBoxPrimitive =>
          primitive.kind === 'textBox' && primitive.storyId === story.id
      );
      if (!textBox) throw new Error('subtitle layout is missing');
      expect(textBox.lines.length).toBeGreaterThan(1);

      const calls: Array<{ text: string; x: number; y: number }> = [];
      const ctx = new Proxy(
        {
          fillText: (text: string, x: number, y: number) => calls.push({ text, x, y }),
        } as Record<string, unknown>,
        {
          get(target, property) {
            if (property in target) return target[property as string];
            return () => undefined;
          },
          set(target, property, value) {
            target[property as string] = value;
            return true;
          },
        }
      ) as unknown as CanvasRenderingContext2D;
      await paintSlide(ctx, { ...frame, background: undefined, primitives: [textBox] });

      const first = textBox.lines[0];
      const painted = calls.filter((call) => call.y === first.baseline);
      expect(painted.length).toBeGreaterThan(1);
      let position = first.start;
      for (const call of painted) {
        const caret = first.caretStops.find((stop) => stop.position === position);
        if (!caret) throw new Error(`caret stop ${position} is missing`);
        expect(call.x).toBe(caret.x);
        position += call.text.length;
      }
      expect(position).toBe(first.end);

      const last = textBox.lines[textBox.lines.length - 1];
      expect(calls.filter((call) => call.y === last.baseline)).toEqual([
        {
          text: last.runs.map((run) => run.text).join(''),
          x: last.runs[0].x,
          y: last.baseline,
        },
      ]);
    } finally {
      source.dispose();
    }
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
  return shapeSnapshotFrom(handle, shapeId);
}

function shapeSnapshotFrom(source: PresentationHandle, shapeId: string) {
  const shape = source.snapshot().slides[0].shapes.find((candidate) => candidate.id === shapeId);
  if (!shape) throw new Error(`shape ${shapeId} was not found`);
  return shape;
}

function firstStory(shapes: Array<{ textStories: StorySnapshot[]; children: unknown[] }>): StorySnapshot {
  for (const shape of shapes) {
    if (shape.textStories[0]) return shape.textStories[0];
  }
  throw new Error('fixture has no text story');
}

function oneCallPerTextRun(
  primitives: SlidePrimitive[]
): Array<{ text: string; x: number; y: number }> {
  const calls: Array<{ text: string; x: number; y: number }> = [];
  for (const primitive of primitives) {
    if (primitive.kind === 'textBox') {
      for (const line of primitive.lines) {
        for (const run of line.runs) calls.push({ text: run.text, x: run.x, y: line.baseline });
      }
    } else if (primitive.kind === 'chart') {
      calls.push(...oneCallPerTextRun(primitive.primitives));
    } else if (primitive.kind === 'placeholder' && primitive.label) {
      calls.push({
        text: primitive.label,
        x: primitive.x + primitive.w / 2,
        y: primitive.y + primitive.h / 2,
      });
    }
  }
  return calls;
}
