import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { PresentationHandle, StorySnapshot, TextBoxPrimitive } from '../index';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
let handle: PresentationHandle;

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  handle = openPresentation(pptx, {
    clientId: 9001,
    fonts: [{ family: 'Liberation Sans', bytes: font }],
  });
});

afterAll(() => handle.dispose());

describe('PPTX wasm boundary', () => {
  test('opens, edits, reflows, hit-tests, and observes a local update', () => {
    const snapshot = handle.snapshot();
    expect(snapshot.slides.length).toBe(3);
    const story = firstStory(snapshot.slides.flatMap((slide) => slide.shapes));
    const insertion = story.length - 1;

    handle.startUpdateObservation();
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

    const event = handle.drainUpdateEvent();
    expect(event?.origin).toBe('local');
    expect(event?.update.length).toBeGreaterThan(0);
    expect(handle.canUndo()).toBe(true);
    expect(handle.undo().applied).toBe(true);
    expect(handle.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('edited'))
    )).toBe(false);
    handle.clearUpdateObservation();
  });
});

function firstStory(shapes: Array<{ textStories: StorySnapshot[]; children: unknown[] }>): StorySnapshot {
  for (const shape of shapes) {
    if (shape.textStories[0]) return shape.textStories[0];
  }
  throw new Error('fixture has no text story');
}
